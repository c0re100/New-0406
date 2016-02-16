var flag=-1; 
var state=0;
var stonenum=0;
var workstate=1;
var steelstate=0;
var passflag=0;
//var cookiename=document.cookie;
var randnum=Math.random();
function doPreload(){
	var the_images = new Array('image/a.gif','image/b.gif','image/c.gif','image/d.gif','image/e.gif','image/f.gif','image/g.gif');
	for(i= 0; i<the_images.length; i++){
   		var an_image = new Image();	
		an_image.src = the_images[i];	
	}
}
doPreload();
//setcookie();
function actstart(){
//	randnum=Math.random();
        setCookie();
        document.shiform.actnum.value=0;
        document.shiform.submit();
}
function checkState(f){ 
	if(f==flag) return true; 
	if(flag==1) alert("現在正在采礦！"); 
	if(flag==2) alert("現在正在煉鐵！"); 
	 
} 
function chechstone(f){
       if(f<100) alert("對不起，你的鐵塊不夠煉鐵！");
}

function setCookie()
{
	var act_state=randnum;
	var cookiename="check_javascript=state:"+act_state;
	document.cookie=cookiename;
//    alert("well");
}

function readCookie()
{   var cookiename=document.cookie;
    var broken_cookie = cookiename.split(":");
    var actstate = broken_cookie[1];
//    alert("Your actstate is: " + actstate+"randnum="+randnum);   

	return true;
}

function gocai(){ 
        if(!readCookie())return;
	hiddenstone();
        if(stonenum>900) alert("請注意鐵塊數量不能大於999！"); 
	flag=1;
        changeImage(caikuang,"image/d.gif");
	caikuang.style.visibility="visible";
	sAniMove(caikuang,5,200,416,250,"goOver()"); 
} 
function hiddenstone(){
	dostone.style.visibility="hidden";
	dosteel.style.visibility="hidden";
	jineng.style.visibility="hidden";
        sjdiv.style.visibility="hidden";
}
function showstone(){
	dostone.style.visibility="visible";
	dosteel.style.visibility="visible";
	jineng.style.visibility="visible";
}

function goOver(){ 
	changeImage(caikuang,"image/b.gif"); 
	setTimeout("startGohome();",6000); 
} 
function startGohome(){ 
	changeImage(caikuang,"image/e.gif"); 
	sAniMove(caikuang,5,200,305,82,"gohomeOver()"); 
} 
function gohomeOver(){ 
	caikuang.style.visibility="hidden"; 
	changeImage(caikuang,"image/d.gif");
        flag=-1; 
	document.shiform.actnum.value=1;
//        alert(passflag);
        document.shiform.passflag.value=passflag; 
	document.shiform.submit();
        showstone();
	workstate=1;
} 
function initT(n1,n2,n3,n4,n5,n6,n7){
	stonenum=n1;
	kstext.innerHTML=n1;
	tktext.innerHTML=n2;
	ksdj.innerHTML=n3;
	tkdj.innerHTML=n4;
        var level=n3+n4+n6-1;
        document.shiform.actnum.value=0;  	
	jingyan.innerHTML="<font title=級別:"+level+"，還需"+((level*(level-1)/2+1)*100-n5)+"點纔能升級>"+n5+"</font>";
	if(n6>0)
		jineng.innerHTML="<a href=# onmousedown='showSjdiv()' class=csstext1>"+n6+"</a>";
	else
    		jineng.innerHTML=n6;
        passflag=n7;
        flag=0;
        workstate=0;
}
function golian(){ 
        if(!readCookie())return;
        flag=2;
        hiddenstone();
        if(steelstate!=0){alert("warning3");return;}
	steelstate=steelstate+1;
//	alert (steelstate);
        changeImage(liantie,"image/c.gif");	
        liantie.style.visibility="visible";
	sAniMove(liantie,7,200,95,278,"goOver1()"); 
} 
function goOver1(){ 
	changeImage(liantie,"image/a.gif"); 
        if(steelstate!=1){alert("warning4");return;}
        steelstate=steelstate+1;
//	alert(steelstate);
	setTimeout("startGohome1();",4000); 
} 
function startGohome1(){ 
	changeImage(liantie,"image/f.gif"); 
        if(steelstate!=2){alert("warning5");return;}
        steelstate=steelstate+1;
//        alert(steelstate);
	sAniMove(liantie,7,200,79,142,"goCaikuang()"); 
} 
function goCaikuang(){ 
	changeImage(liantie,"image/g.gif"); 
        if(steelstate!=3){alert("warning6");return;}
        steelstate=steelstate+1;
//        alert(steelstate);
	sAniMove(liantie,7,200,281,79,"gohomeOver1()"); 
} 
function gohomeOver1(){ 
	liantie.style.visibility="hidden";
        changeImage(liantie,"image/c.gif");
	flag=-1;
        if(steelstate=4){
	document.shiform.actnum.value=(steelstate/2);
	document.shiform.passflag.value=passflag;
	steelstate=0;
        document.shiform.submit();}
        showstone();
//        reset();
//	document.shiform.actnum.value=2;
//	document.shiform.submit();
//        steelstate=0;
//        document.shiform.actnum.value=0;
//	document.shiform.submit();
} 
function reset(){
	document.shiform.actnum.value=0;
	document.shiform.submit();
}
function showSjdiv(){
	if(!checkState(0)) return; 
	if(sjdiv.style.visibility=="hidden")
		sjdiv.style.visibility="visible";
	else
		sjdiv.style.visibility="hidden"
}
function getJnd(){
	if(document.form1.sjbutton[0].checked==true){
		document.shiform.actnum.value=3;
	}
	if(document.form1.sjbutton[1].checked==true){
		document.shiform.actnum.value=4;
	}
	document.shiform.submit();
	sjdiv.style.visibility="hidden";
}

function sAniMove(o,t,m,nx,ny,om){
	aniobj=o;
	aniobj.anioldx=aniobj.style.pixelLeft;
	aniobj.anioldy=aniobj.style.pixelTop;
	aniobj.anixx=nx;aniyy=ny;
	aniobj.overMethod=om;
	aniobj.anistepx=0;aniobj.anistepy=0;
	aniobj.anittm=0;aniobj.anitm=0;
	with(aniobj){
		anitm=Math.floor(Math.sqrt((anixx-anioldx)*(anixx-anioldx)+(aniyy-anioldy)*(aniyy-anioldy))/t);
//	alert(anitm);	
	anistepx=(anixx-anioldx)/anitm;
		anistepy=(aniyy-anioldy)/anitm;
	}
	aniobj.anitime=setInterval("aniMove()",m);
}
function aniMove(){
	with(aniobj){
		if(anittm>=anitm){
			clearInterval(anitime);
			style.pixelLeft=anixx; style.pixelTop=aniyy;
			if(overMethod!=null) eval(overMethod);
		}else{
			style.pixelLeft=anioldx+anistepx; anioldx+=anistepx;
			style.pixelTop=anioldy+anistepy; anioldy+=anistepy;
			anittm++;
		}
	}
}

function changeImage(img, newSource){
	img.src = newSource;
}
