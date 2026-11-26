// Builds app/index.html from the shared page (src/retzef.html): swaps CDN links for bundled local copies so the app works offline.
const fs=require('fs'),path=require('path');
const root=path.join(__dirname,'..'),out=path.join(root,'app');
fs.rmSync(out,{recursive:true,force:true});fs.mkdirSync(out,{recursive:true});
fs.cpSync(path.join(root,'vendor'),path.join(out,'vendor'),{recursive:true});
let html=fs.readFileSync(path.join(root,'src','retzef.html'),'utf8');
const faces=[];
for(const [fam,file,weights] of [['Rubik','rubik',[500,700,800]],['Assistant','assistant',[400,600,700]]])
  for(const w of weights)for(const [sub,range] of [['hebrew','U+0307-0308,U+0590-05FF,U+200C-2010,U+20AA,U+25CC,U+FB1D-FB4F'],['latin','U+0000-00FF,U+0131,U+0152-0153,U+02BB-02BC,U+02C6,U+02DA,U+02DC,U+2000-206F,U+20AC,U+2122,U+2191,U+2193,U+2212,U+2215,U+FEFF,U+FFFD']])
    faces.push(`@font-face{font-family:"${fam}";font-weight:${w};font-style:normal;font-display:swap;src:url(vendor/fonts/${file}-${sub}-${w}-normal.woff2) format("woff2");unicode-range:${range}}`);
const rep=(a,b)=>{if(!html.includes(a))throw new Error('prepare: not found: '+a.slice(0,60));html=html.split(a).join(b)};
rep('<link rel="preconnect" href="https://fonts.googleapis.com">\n','');
rep('<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Rubik:wght@500;700;800&family=Assistant:wght@400;600;700&display=swap">','<style>'+faces.join('\n')+'</style>');
rep('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js','vendor/pdf.min.js');
rep('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js','vendor/pdf.worker.min.js');
rep('https://cdn.jsdelivr.net/npm/pptx-preview@1.0.7/dist/pptx-preview.umd.js','vendor/pptx-preview.umd.js');
rep('https://cdnjs.cloudflare.com/ajax/libs/jszip/3.10.1/jszip.min.js','vendor/jszip.min.js');
html='<!doctype html>\n<html lang="he" dir="rtl">\n<head>\n<meta charset="utf-8">\n<meta name="viewport" content="width=device-width, initial-scale=1">\n'+
  '<meta http-equiv="Content-Security-Policy" content="default-src \'self\' data: blob:; script-src \'self\' \'unsafe-inline\'; style-src \'self\' \'unsafe-inline\'; img-src * data: blob:; media-src * data: blob:; frame-src *; connect-src \'self\' data: blob:">\n'+
  html+'\n</html>\n';
fs.writeFileSync(path.join(out,'index.html'),html);
console.log('app/index.html ready');
