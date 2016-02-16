
//Pre-load your image below!
Image0=new Image();
Image0.src="pic/hua1.gif";
Image1=new Image();
Image1.src="pic/hua1.gif";
Image2=new Image();
Image2.src="pic/hua2.gif";
Image3=new Image();
Image3.src="pic/hua.gif";
Image4=new Image();
Image4.src="pic/hua1.gif";
Image5=new Image();
Image5.src="pic/hua2.gif"; 

grphcs=new Array(6)
grphcs[0]="pic/hua1.gif"
grphcs[1]="pic/hua1.gif"
grphcs[2]="pic/hua2.gif"
grphcs[3]="pic/hua.gif"
grphcs[4]="pic/hua1.gif"
grphcs[5]="pic/hua2.gif"



Amount=10; //Smoothness depends on image file size, the smaller the size the more you can use!
Ypos=new Array();
Xpos=new Array();
Speed=new Array();
Step=new Array();
Cstep=new Array();
ns=(document.layers)?1:0;

if (ns){
for (i = 0; i < Amount; i++){
var P=Math.floor(Math.random()*grphcs.length);
rndPic=grphcs[P];
document.write("<LAYER NAME='sn"+i+"' LEFT=0 TOP=0><img src="+rndPic+"></LAYER>");
}
}
else{
document.write('<div style="position:absolute;top:0px;left:0px"><div style="position:relative">');
for (i = 0; i < Amount; i++){
var P=Math.floor(Math.random()*grphcs.length);
rndPic=grphcs[P];
document.write('<img id="si" src="'+rndPic+'" style="position:absolute;top:0px;left:0px">');
}
document.write('</div></div>');
}
WinHeight=(document.layers)?window.innerHeight:window.document.body.clientHeight;
WinWidth=(document.layers)?window.innerWidth:window.document.body.clientWidth;
for (i=0; i < Amount; i++){                                                                
 Ypos[i] = Math.round(Math.random()*WinHeight);
 Xpos[i] = Math.round(Math.random()*WinWidth);
 Speed[i]= Math.random()*5+3;
 Cstep[i]=0;
 Step[i]=Math.random()*0.1+0.05;
}
function fall(){
var WinHeight=(document.layers)?window.innerHeight:window.document.body.clientHeight;
var WinWidth=(document.layers)?window.innerWidth:window.document.body.clientWidth;
var hscrll=(document.layers)?window.pageYOffset:document.body.scrollTop;
var wscrll=(document.layers)?window.pageXOffset:document.body.scrollLeft;
for (i=0; i < Amount; i++){
sy = Speed[i]*Math.sin(90*Math.PI/180);
sx = Speed[i]*Math.cos(Cstep[i]);
Ypos[i]+=sy;
Xpos[i]+=sx; 
if (Ypos[i] > WinHeight){
Ypos[i]=-60;
Xpos[i]=Math.round(Math.random()*WinWidth);
Speed[i]=Math.random()*5+3;
}
if (ns){
document.layers['sn'+i].left=Xpos[i];
document.layers['sn'+i].top=Ypos[i]+hscrll;
}
else{
si[i].style.pixelLeft=Xpos[i];
si[i].style.pixelTop=Ypos[i]+hscrll;
} 
Cstep[i]+=Step[i];
}
setTimeout('fall()',50);
}
