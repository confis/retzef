// Bridge between the page and the desktop features (PowerPoint control, screens). Exposes a small, fixed API only.
const {contextBridge,ipcRenderer}=require('electron');
const on=(ch,fn)=>{const h=(_e,d)=>fn(d);ipcRenderer.on(ch,h);return()=>ipcRenderer.removeListener(ch,h)};
contextBridge.exposeInMainWorld('retzefDesktop',{
  platform:process.platform,
  pptCheck:()=>ipcRenderer.invoke('ppt:check'),
  pptPrepare:(id,ext,buffer)=>ipcRenderer.invoke('ppt:prepare',{id,ext,buffer}),
  pptHasFile:(id,ext)=>ipcRenderer.invoke('ppt:has',{id,ext}),
  pptStart:(opts)=>ipcRenderer.invoke('ppt:start',opts),
  pptNext:()=>ipcRenderer.invoke('ppt:cmd','next'),
  pptPrev:()=>ipcRenderer.invoke('ppt:cmd','prev'),
  pptStop:()=>ipcRenderer.invoke('ppt:stop'),
  pptRemove:(id,ext)=>ipcRenderer.invoke('ppt:remove',{id,ext}),
  screens:()=>ipcRenderer.invoke('screens'),
  setFullscreen:(on)=>ipcRenderer.invoke('fullscreen',on),
  onPptState:(fn)=>on('ppt:state',fn),
  onPptEnded:(fn)=>on('ppt:ended',fn)
});
