let _xlsxLoaded=false,_xlsxCbs=[];
function ensureXLSX(fn){
  if(typeof XLSX!=='undefined'){fn();return;}
  _xlsxCbs.push(fn);
  if(_xlsxLoaded)return;
  _xlsxLoaded=true;
  const s=document.createElement('script');
  s.src='https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';
  s.onload=function(){_xlsxCbs.forEach(c=>c());_xlsxCbs=[];};
  document.head.appendChild(s);
}
