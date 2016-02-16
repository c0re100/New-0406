var isns4up = (document.layers) ? 1 : 0;
var isie4up = (document.all) ? 1 : 0;
var dyy, xp0,xp1, yp0,yp1;
var amm, stxx, styy;
var d_width = 800;

if (isns4up) d_width = self.innerWidth-150;
else if (isie4up) d_width = document.body.clientWidth-150;

dyy = 0;
xp1 = Math.random()*(d_width);
yp1 = 50
xp0 = Math.random()*(d_width);
yp0 = 50
amm = Math.random()*20;
stxx = 0.02 + Math.random()/10;
styy = 1.7 + Math.random();

if (isns4up) {
document.write("<layer name=\"a0\" left=\"15\" top=\"15\" visibility=\"show\"><img src='data/p26.gif' border=\"0\"></layer>");
document.write("<layer name=\"a1\" left=\"15\" top=\"15\" visibility=\"show\"><img src='data/123.gif' border=\"0\"></layer>");
} else if (isie4up) {
document.write("<div id=\"a0\" style=\"POSITION: absolute; Z-INDEX: 0; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><img src='data/p26.gif' border=\"0\"></div>");
document.write("<div id=\"a1\" style=\"POSITION: absolute; Z-INDEX: 1; VISIBILITY: visible; TOP: 15px; LEFT: 15px;\"><img src='data/123.gif' border=\"0\"></div>");
}

function aNS() {
if (xp0<50) {xp0=d_width;yp0=50;}
if (xp1>d_width) {xp1=50;yp1=50;}
styy = 0.02 + Math.random()/10;
stxx = 1.7 + Math.random();
xp0 -= stxx;
xp1 += stxx;
dyy += styy;
document.layers["a0"].top = yp0+ amm*Math.sin(dyy);
document.layers["a0"].left = xp0 ;
document.layers["a1"].top = yp1+ amm*Math.sin(dyy);
document.layers["a1"].left = xp1 ;
setTimeout("aNS()", 20);
}

function aIE() {
if (xp0<50) {xp0=d_width;yp0=50;}
if (xp1>d_width) {xp1=50;yp1=50;}
styy = 0.02 + Math.random()/10;
stxx = 1.7 + Math.random();
xp0 -= stxx;
xp1 += stxx;
dyy += styy;
document.all["a0"].style.pixelTop = yp0+ amm*Math.sin(dyy);
document.all["a0"].style.pixelLeft = xp0 ;
document.all["a1"].style.pixelTop = yp1+ amm*Math.sin(dyy);
document.all["a1"].style.pixelLeft = xp1 ;
setTimeout("aIE()", 20);
}

if (isns4up) {
aNS();
} else if (isie4up) {
aIE();
} 