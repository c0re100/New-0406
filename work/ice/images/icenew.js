var flag=0;
var state=0;
var pn=0;
var icenum=0;
function checkState(f){
	if(f==flag) return true;
	if(flag==0) alert("現在正在休息！");
	if(flag==1) alert("現在正在采冰！");
	if(flag==2) alert("現在需要化冰！");
	if(flag==3) alert("現在正在化冰！");
}
function go(){
	if(!checkState(0)) return;
	flag=1;
	changeImage(person,"images/go1.gif");
	sAniMove(person,7,120,410,350,0,"goOver()");
}
function goOver(){
	changeImage(person,"images/go2.gif");
	setTimeout("changeImage(person,'images/to0.gif');flag=2;",2000);
}
function to(){
	if(!checkState(2)) return;
	flag=3;
	changeImage(person,"images/to1.gif");
	sAniMove(person,7,100,200,315,0,"toOver()");
}
function toOver(){
	person.style.visibility="hidden";
	huo.style.visibility="visible";
	yan.style.visibility="visible";
	setTimeout("oOver()",2000);
}
function oOver(){
	huo.style.visibility="hidden";
	yan.style.visibility="hidden";
	changeImage(person,"images/go0.gif");
	person.style.visibility="visible";
	flag=0;
	sState();
}	
function sState(){
	document.iceform.needit.value=state;
	document.iceform.submit();
	state++;
	if(state > 4) state=0;

	pn=document.iceform.itneed.value;
	if(pn>4) pn=0;
	changeImage(icenum,"images/tong"+pn+".gif");
}
	
function sAniMove(o,t,m,nx,ny,type,om){
	document.aniobj=o;
	document.anioldx=document.aniobj.style.pixelLeft;
	document.anioldy=document.aniobj.style.pixelTop;
	document.anixx=nx;document.aniyy=ny;
	document.overMethod=om;
	if(type!=1){
		var zf=1;
		if((document.anixx-document.anioldx)<0) zf=-1;
		document.anitm=Math.floor(zf*(document.anixx-document.anioldx)/t);
		document.anistepx=zf*t;
		document.anistepy=0;
		if((document.anixx-document.anioldx)!=0)
			document.anistepy=zf*t*(document.aniyy-document.anioldy)/(document.anixx-document.anioldx);
	}else{
		document.anitm=Math.floor(1000*t/m);
		document.anistepx=(document.anixx-anioldx)/document.anitm;
		document.anistepy=(document.anixx-anioldx)/document.anitm;
	}
	document.anittm=0;
	document.anitime=setInterval("aniMove()",m);
}
function aniMove(){
	if(document.anittm>document.anitm){
		clearInterval(document.anitime);
		if(document.overMethod!=null) eval(document.overMethod);
	}else{
		document.aniobj.style.pixelLeft=document.anioldx+document.anistepx;
		document.anioldx+=document.anistepx;
		document.aniobj.style.pixelTop=document.anioldy+document.anistepy;
		document.anioldy+=document.anistepy;
		document.anittm++;
	}
}
function changeImage(img, newSource){
	img.src = newSource;
}
