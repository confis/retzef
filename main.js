// Retzef desktop – Electron main process.
// The page (app/index.html) is the same player as the web version; this process adds:
//  • PowerPoint playback through the installed PowerPoint (ppt-host.ps1, COM automation)
//  • the audience window on the second screen, and embedding of any web page.
const {app,BrowserWindow,ipcMain,screen,shell,Menu,session}=require('electron');
const path=require('path'),fs=require('fs'),{spawn}=require('child_process');

let win=null,host=null,hostBuf='',seqId=0;
const pending=new Map();
let poll=null,polling=false,lastState=null,showSession=null;

/* ---------- PowerPoint host process ---------- */
const useFake=()=>process.platform!=='win32'||!!process.env.RETZEF_FAKE_PPT;
function hostScript(){return path.join(__dirname,'ppt-host.ps1').replace('app.asar'+path.sep,'app.asar.unpacked'+path.sep)}
function startHost(){
  if(host)return host;
  host=useFake()
    ?spawn(process.execPath,[path.join(__dirname,'fake-ppt.js')],{env:{...process.env,ELECTRON_RUN_AS_NODE:'1'}})
    :spawn('powershell.exe',['-NoProfile','-NonInteractive','-ExecutionPolicy','Bypass','-Sta','-File',hostScript()],{windowsHide:true});
  hostBuf='';
  host.stdout.setEncoding('utf8');
  host.stdout.on('data',d=>{
    hostBuf+=d;let i;
    while((i=hostBuf.indexOf('\n'))>=0){
      const line=hostBuf.slice(0,i).trim();hostBuf=hostBuf.slice(i+1);if(!line)continue;
      let m;try{m=JSON.parse(line)}catch(e){continue}
      const p=pending.get(m.id);if(p){pending.delete(m.id);clearTimeout(p.t);p.res(m)}
    }
  });
  host.stderr.on('data',d=>console.error('[ppt-host]',String(d).trim()));
  host.on('error',e=>console.error('[ppt-host] failed to start',e.message));
  host.on('exit',()=>{host=null;for(const p of pending.values()){clearTimeout(p.t);p.res({ok:false,error:'PowerPoint host stopped'})}pending.clear();if(showSession)endShow(false)});
  return host;
}
function hostCmd(cmd,extra={},timeout=90000){
  return new Promise(res=>{
    let h;try{h=startHost()}catch(e){return res({ok:false,error:e.message})}
    const id=++seqId;const t=setTimeout(()=>{pending.delete(id);res({ok:false,error:'PowerPoint did not answer in time'})},timeout);
    pending.set(id,{res,t});
    try{h.stdin.write(JSON.stringify({id,cmd,...extra})+'\n')}catch(e){clearTimeout(t);pending.delete(id);res({ok:false,error:e.message})}
  });
}

/* ---------- files: PowerPoint needs a real file on disk ---------- */
const presDir=()=>{const d=path.join(app.getPath('userData'),'presentations');fs.mkdirSync(d,{recursive:true});return d};
const safeId=s=>String(s).replace(/[^\w-]/g,'');
const presPath=(id,ext)=>path.join(presDir(),safeId(id)+'.'+(ext==='ppsx'?'ppsx':'pptx'));

/* ---------- screens ---------- */
function audienceDisplay(){
  const all=screen.getAllDisplays(),prim=screen.getPrimaryDisplay();
  if(all.length<2)return null;
  const winD=win?screen.getDisplayMatching(win.getBounds()):prim;
  return all.find(d=>d.id!==winD.id)||null;
}
async function physicalScreenFor(d){
  if(!d)return null;
  const r=await hostCmd('screens',{},20000);if(!r.ok||!r.screens||!r.screens.length)return null;
  const tx=d.bounds.x*d.scaleFactor,ty=d.bounds.y*d.scaleFactor;
  let best=null,bd=Infinity;
  const isPrimary=d.id===screen.getPrimaryDisplay().id;
  const cands=r.screens.filter(s=>!!s.primary===isPrimary);
  for(const s of (cands.length?cands:r.screens)){const dist=Math.abs(s.x-tx)+Math.abs(s.y-ty);if(dist<bd){bd=dist;best=s}}
  return best;
}

/* ---------- slide-show session: poll PowerPoint and report to the page ---------- */
function startPolling(){
  stopPolling();
  poll=setInterval(async()=>{
    if(polling||!showSession)return;polling=true;
    const r=await hostCmd('state',{},10000);polling=false;
    if(!showSession||!r.ok)return;
    if(r.active){
      lastState=r;send('ppt:state',{slide:r.slide,count:r.count,click:r.click,clicks:r.clicks,state:r.state});
      if(r.state===5){await hostCmd('stop');endShow(true)}   // reached PowerPoint's end screen
    }else if(showSession.seenActive){
      endShow(!!(lastState&&lastState.slide>=lastState.count)); // closed in PowerPoint (Esc) or ended without end screen
    }
    if(r.active)showSession.seenActive=true;
  },250);
}
function stopPolling(){if(poll)clearInterval(poll);poll=null;polling=false}
function endShow(completed){
  if(!showSession)return;
  const id=showSession.id;showSession=null;lastState=null;stopPolling();
  send('ppt:ended',{id,completed});
  if(win){win.show();win.focus()}
}
function send(ch,data){if(win&&!win.isDestroyed())win.webContents.send(ch,data)}
function refocusMain(){
  if(!win)return;
  // take keyboard focus back so the clicker keeps talking to Retzef (it forwards to PowerPoint)
  win.setAlwaysOnTop(true);win.show();win.focus();setTimeout(()=>{if(win)win.setAlwaysOnTop(false)},400);
}

/* ---------- IPC ---------- */
ipcMain.handle('ppt:check',async()=>{const r=await hostCmd('check',{},60000);return {ok:!!r.ok,version:r.version||null,error:r.error||null,simulated:useFake()}});
ipcMain.handle('ppt:has',(_e,{id,ext})=>fs.existsSync(presPath(id,ext)));
ipcMain.handle('ppt:prepare',async(_e,{id,ext,buffer})=>{
  const p=presPath(id,ext);
  if(buffer)fs.writeFileSync(p,Buffer.from(buffer));
  if(!fs.existsSync(p))return {ok:false,error:'missing file'};
  const r=await hostCmd('open',{path:p});
  return {ok:!!r.ok,count:r.count,error:r.error};
});
ipcMain.handle('ppt:remove',async(_e,{id,ext})=>{const p=presPath(id,ext);if(host)await hostCmd('close',{path:p},10000);try{fs.unlinkSync(p)}catch(e){}return true});
ipcMain.handle('ppt:start',async(_e,o)=>{
  const p=presPath(o.id,o.ext);
  if(showSession){showSession=null;stopPolling()}
  const d=audienceDisplay();
  const scr=d?await physicalScreenFor(d):null;
  const r=await hostCmd('start',{path:p,at:o.at||1,screen:scr});
  if(!r.ok)return r;
  showSession={id:o.id,seenActive:false};startPolling();
  if(d)setTimeout(refocusMain,700);
  return {ok:true,count:r.count,dual:!!d};
});
ipcMain.handle('ppt:cmd',async(_e,c)=>{if(!['next','prev'].includes(c))return {ok:false};return hostCmd(c,{},15000)});
ipcMain.handle('ppt:stop',async()=>{const had=!!showSession;showSession=null;stopPolling();lastState=null;if(host)await hostCmd('stop',{},15000);if(had&&win)win.focus();return true});
ipcMain.handle('screens',()=>({count:screen.getAllDisplays().length}));
ipcMain.handle('fullscreen',(_e,on)=>{if(win)win.setFullScreen(!!on);return true});

/* ---------- windows ---------- */
function createWindow(){
  win=new BrowserWindow({
    width:1360,height:860,minWidth:900,minHeight:600,backgroundColor:'#0e1219',title:'Retzef',
    autoHideMenuBar:true,
    webPreferences:{preload:path.join(__dirname,'preload.js'),contextIsolation:true,nodeIntegration:false,sandbox:true}
  });
  win.loadFile(path.join(__dirname,'app','index.html'));
  win.webContents.setWindowOpenHandler(({url,frameName})=>{
    if(frameName==='retzef-audience'){
      const d=audienceDisplay();
      const b=d?d.bounds:{x:undefined,y:undefined,width:1280,height:720};
      return {action:'allow',overrideBrowserWindowOptions:{x:b.x,y:b.y,width:b.width,height:b.height,fullscreen:!!d,autoHideMenuBar:true,backgroundColor:'#000000',title:'Retzef – מסך קהל'}};
    }
    if(/^https?:/i.test(url))shell.openExternal(url);
    return {action:'deny'};
  });
  win.webContents.on('before-input-event',(_e,i)=>{if(i.type==='keyDown'&&i.key==='F12')win.webContents.toggleDevTools()});
  win.on('closed',()=>{win=null});
}

app.whenReady().then(()=>{
  Menu.setApplicationMenu(null);
  // let web-page slides embed sites that normally refuse to be framed
  session.defaultSession.webRequest.onHeadersReceived((d,cb)=>{
    if(d.resourceType!=='subFrame')return cb({});
    const h={...d.responseHeaders};
    for(const k of Object.keys(h)){
      const lk=k.toLowerCase();
      if(lk==='x-frame-options')delete h[k];
      if(lk==='content-security-policy')h[k]=h[k].map(v=>v.replace(/frame-ancestors[^;]*;?/gi,''));
    }
    cb({responseHeaders:h});
  });
  createWindow();
});
let quitting=false;
app.on('before-quit',async e=>{
  if(quitting||!host)return;
  e.preventDefault();quitting=true;
  await hostCmd('quit',{},8000);try{host&&host.kill()}catch(err){}
  app.quit();
});
app.on('window-all-closed',()=>app.quit());
