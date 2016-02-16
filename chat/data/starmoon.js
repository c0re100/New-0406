var no = 2;

var ns4up = (document.layers) ? 1 : 0;
var ie4up = (document.all) ? 1 : 0;

var dx, xp, yp;
var am, stx, sty;
var i, x_len , y_len , doc_width = 800, doc_height = 600;

if (ns4up) {
doc_width = self.innerWidth;
if (self.outerHeight>self.innerHeight)
doc_height = self.outerHeight;
else
doc_height = self.innerHeight;
} else if (ie4up) {
doc_width = document.body.clientWidth;
if (document.body.scrollHeight>document.body.clientHeight)
doc_height = document.body.scrollHeight;
else
doc_height = document.body.clientHeight;
}

dx = new Array();
xp = new Array();
yp = new Array();
am = new Array();
stx = new Array();
sty = new Array();

for (i = 0; i < no; ++ i) {
dx[i] = 0;
xp[i] = Math.random()*(doc_width-100);
yp[i] = Math.random()*(doc_height-30);
am[i] = Math.random()*20;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 1.7 + Math.random();
}
x_len = 28;
y_len = 28;
if (ns4up) {
document.write("<layer name='dot0' left='15' top='15' visibility='show'><img src='data/star2.gif' border='0'></layer>");
document.write("<layer name='dot1' left='15' top='15' visibility='show'><img src='data/star2.gif' border='0'></layer>");
document.write("<layer name='dot2' left='15' top='15' visibility='show'><img alt='還記得:你曾說過陪我一起看星星' src='data/12.gif' border='0'></layer>");
} else if (ie4up) {
document.write("<div id='dot0' style='POSITION: absolute; Z-INDEX: 2; VISIBILITY: visible; TOP: 15px; LEFT: "+y_len+"px;'><img src='data/star2.gif' border='0'></div>");
document.write("<div id='dot1' style='POSITION: absolute; Z-INDEX: 1; VISIBILITY: visible; TOP: 15px; LEFT: "+y_len+"px;'><img src='data/star2.gif' border='0'></div>");
document.write("<div id='dot2' style='POSITION: absolute; Z-INDEX: 0; VISIBILITY: visible; TOP: 15px; LEFT: "+y_len+"px;'><img alt='還記得:你曾說過陪我一起看星星' src='data/12.gif' border='0'></div>");
}

function plumNS() {
if (self.outerHeight>self.innerHeight)
doc_height = self.outerHeight;
else
doc_height = self.innerHeight;
doc_width = self.innerWidth;
document.layers["dot2"].top = doc_height-self.innerHeight+10;
document.layers["dot2"].left = doc_width-200;
for (i = 0; i < no; ++ i) {
yp[i] += sty[i];
if (yp[i] > (doc_height-80)) {
xp[i] = Math.random()*(doc_width-am[i]-100);
yp[i] = doc_height-self.innerHeight+20;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 1.7 + Math.random();
}
dx[i] += stx[i];
document.layers["dot"+i].top = yp[i];
document.layers["dot"+i].left = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("plumNS()", 50);
}

function plumIE() {
if (document.body.scrollHeight>document.body.clientHeight)
doc_height = document.body.scrollHeight;
else
doc_height = document.body.clientHeight;
doc_width = document.body.clientWidth;
document.all["dot2"].style.pixelTop = doc_height-document.body.clientHeight+10;
document.all["dot2"].style.pixelLeft = doc_width-200;
for (i = 0; i < no; ++ i) {
yp[i] += sty[i];
if (yp[i] > (doc_height-80) ){
xp[i] = Math.random()*(doc_width-am[i]-100);
yp[i] = doc_height-document.body.clientHeight+20;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 1.7 + Math.random();
}
dx[i] += stx[i];
document.all["dot"+i].style.pixelTop = yp[i];
document.all["dot"+i].style.pixelLeft = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("plumIE()", 50);
}

if (ns4up) {
plumNS();
} else if (ie4up) {
plumIE();
} 