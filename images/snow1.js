<!--
SmallStars = 2; 
LargeStars = 2;

SmallYpos = new Array();
SmallXpos = new Array();
LargeYpos = new Array();
LargeXpos = new Array();
Smallspeed= new Array();
Largespeed= new Array();
ns=(document.layers)?1:0;
if (ns){
for (i = 0; i < SmallStars; i++)
{document.write("<LAYER NAME='sn"+i+"' LEFT=0 TOP=0 BGCOLOR='#FFFFF0' CLIP='0,0,1,1'></LAYER>")}
for (i = 0; i < LargeStars; i++)
{document.write("<LAYER NAME='ln"+i+"' LEFT=0 TOP=0 BGCOLOR='#FFFFFF' CLIP='0,0,2,2'></LAYER>")}
}
else{
document.write('<div style="position:absolute;top:0px;left:0px">');
document.write('<div style="position:relative">');
for (i = 0; i < SmallStars; i++)
{document.write('<div id="si" style="position:absolute;top:0;left:0;width:1px;height:1px;font-size:8px;font-color:fffff0">*</div>')}
document.write('</div>');
document.write('</div>');
document.write('<div style="position:absolute;top:0px;left:0px">');
document.write('<div style="position:relative">');
for (i = 0; i < LargeStars; i++) 
{document.write('<div id="li" style="position:absolute;top:0;left:0;width:2px;height:2px;font-size:12px;font-color:ffffff">*</div>')}
document.write('</div>');
document.write('</div>');
}
WinHeight=(document.layers)?window.innerHeight:window.document.body.clientHeight;
WinWidth=(document.layers)?window.innerWidth:window.document.body.clientWidth;
//Inital placement!
for (i=0; i < SmallStars; i++) 
 {                                                                
 SmallYpos[i] = Math.round(Math.random()*WinHeight);
 SmallXpos[i] = Math.round(Math.random()*WinWidth);
 Smallspeed[i]= Math.random()*5+1;
}
for (i=0; i < LargeStars; i++) 
 {                                                                
 LargeYpos[i] = Math.round(Math.random()*WinHeight);
 LargeXpos[i] = Math.round(Math.random()*WinWidth);
 Largespeed[i]= Math.random()*10+5;
}
function fly(){
var WinHeight=(document.layers)?window.innerHeight:window.document.body.clientHeight;
var WinWidth=(document.layers)?window.innerWidth:window.document.body.clientWidth;
var hscrll=(document.layers)?window.pageYOffset:document.body.scrollTop;
var wscrll=(document.layers)?window.pageXOffset:document.body.scrollLeft;

for (i=0; i < LargeStars; i++)
{
 LargeXpos[i]-=Largespeed[i];
 if (LargeXpos[i] < -10)
  {
  LargeXpos[i]=WinWidth;
  LargeYpos[i]=Math.round(Math.random()*WinHeight);
  Largespeed[i]=Math.random()*10+5;
  }
 if (ns){
 document.layers['ln'+i].left=LargeXpos[i];
 document.layers['ln'+i].top=LargeYpos[i]+hscrll;
 }
 else{
 li[i].style.pixelLeft=LargeXpos[i];
 li[i].style.pixelTop=LargeYpos[i]+hscrll;
 }
}

for (i=0; i < SmallStars; i++)
{
 SmallXpos[i]-=Smallspeed[i];
 if (SmallXpos[i] < -10)
  {
  SmallXpos[i]=WinWidth;
  SmallYpos[i]=Math.round(Math.random()*WinHeight);
  Smallspeed[i]=Math.random()*5+1;
  }
 if (ns){
 document.layers['sn'+i].left=SmallXpos[i];
 document.layers['sn'+i].top=SmallYpos[i]+hscrll;
 }
 else{
 si[i].style.pixelLeft=SmallXpos[i];
 si[i].style.pixelTop=SmallYpos[i]+hscrll;
 } 
}
setTimeout('fly()',10);
}
fly();
//-->