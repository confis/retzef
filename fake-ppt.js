// Stand-in for ppt-host.ps1 on machines without PowerPoint (development / tests on macOS & Linux).
// Same line-based JSON protocol; simulates a 5-slide deck where slide 2 has 2 click animations.
const rl=require('readline').createInterface({input:process.stdin});
const pres=new Map();let show=null;
const clicksOf=s=>s===2?2:0;
const out=o=>process.stdout.write(JSON.stringify(o)+'\n');
rl.on('line',line=>{
  let m;try{m=JSON.parse(line)}catch(e){return}
  const r={id:m.id,ok:true};
  switch(m.cmd){
    case 'check':r.version='16.0 (simulated)';break;
    case 'screens':r.screens=[{name:'\\\\.\\DISPLAY1',primary:true,x:0,y:0,w:1920,h:1080}];break;
    case 'open':pres.set(m.path,5);r.count=5;break;
    case 'start':{const n=pres.get(m.path)||5;show={path:m.path,count:n,slide:m.at==='last'?n:Math.max(1,+m.at||1),click:0,done:false};if(m.at==='last')show.click=clicksOf(n);r.count=n;break}
    case 'next':if(show){if(show.click<clicksOf(show.slide))show.click++;else if(show.slide<show.count){show.slide++;show.click=0}else show.done=true}break;
    case 'prev':if(show){if(show.done)show.done=false;else if(show.click>0)show.click--;else if(show.slide>1){show.slide--;show.click=clicksOf(show.slide)}}break;
    case 'state':if(show){r.active=true;r.state=show.done?5:1;r.count=show.count;if(!show.done){r.slide=show.slide;r.click=show.click;r.clicks=clicksOf(show.slide)}r.path=show.path}else r.active=false;break;
    case 'stop':show=null;break;
    case 'close':pres.delete(m.path);break;
    case 'quit':out(r);process.exit(0);
    default:r.ok=false;r.error='unknown command';
  }
  out(r);
});
