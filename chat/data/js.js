var no = 5;

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
xp[i] = Math.random()*(doc_width-70);
yp[i] = Math.random()*(doc_height-30);
am[i] = Math.random()*20;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 1.7 + Math.random();
x_len = 28
y_len = 28
if (ns4up) {
if (i == 0) {
document.write("<layer name=\"dot"+ i +"\" left=\"15\" top=\"15\" visibility=\"show\"><img src='data/h"+i+".gif' border=\"0\"></a></layer>");
} else {
document.write("<layer name=\"dot"+ i +"\" left=\"15\" top=\"15\" visibility=\"show\"><img src='data/h"+i+".gif' border=\"0\"></layer>");
}
} else if (ie4up) {
if (i == 0) {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: "+y_len+"px;\"><img src='data/h"+i+".gif' border=\"0\"></a></div>");
} else {
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: absolute; Z-INDEX: "+ i +"; VISIBILITY: visible; TOP: 15px; LEFT: "+y_len+"px;\"><img src='data/h"+i+".gif' border=\"0\"></div>");
}
}
}

function plumNS() {
for (i = 0; i < no; ++ i) {
yp[i] -= sty[i];
if (yp[i] < (doc_height-self.innerHeight+5)) {
xp[i] = Math.random()*(doc_width-am[i]-70);
yp[i] = doc_height-40;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 1.7 + Math.random();
doc_width = self.innerWidth;
if (self.outerHeight>self.innerHeight)
doc_height = self.outerHeight;
else
doc_height = self.innerHeight;
}
dx[i] += stx[i];
document.layers["dot"+i].top = yp[i];
document.layers["dot"+i].left = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("plumNS()", 20);
}

function plumIE() {
for (i = 0; i < no; ++ i) {
yp[i] -= sty[i];
if (yp[i] < (doc_height-document.body.clientHeight+5)) {
xp[i] = Math.random()*(doc_width-am[i]-70);
yp[i] = doc_height-40;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 1.7 + Math.random();
doc_width = document.body.clientWidth;
if (document.body.scrollHeight>document.body.clientHeight)
doc_height = document.body.scrollHeight;
else
doc_height = document.body.clientHeight;
}
dx[i] += stx[i];
document.all["dot"+i].style.pixelTop = yp[i];
document.all["dot"+i].style.pixelLeft = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("plumIE()", 20);
}

if (ns4up) {
plumNS();
} else if (ie4up) {
plumIE();
} 