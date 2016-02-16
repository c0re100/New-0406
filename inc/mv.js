<!-- modify : WWW.CN.TOM.COM-------->
<!-- date :   2000/09/30 -->
<!-- Begin
var ns4up = (document.layers) ? 1 : 0;  // browser sniffer
var ie4up = (document.all) ? 1 : 0;
var no = 1;                             // pic number
var speed = 30;                         // smaller number moves the pic faster
var snowflake ="http://127.0.0.1/xpl/inc/sail.swf";             //pic image
var filen = "http://www.hdt.net.cn/adsl/adslyh.htm"; 	  // link web file
var dx, xp, yp;                         // coordinate and position variables
var am, stx, sty;                       // amplitude and step variables
var i, doc_width = 400, doc_height = 600;
if (ns4up) {
doc_width = self.innerWidth;
doc_height = self.innerHeight;
} else if (ie4up) {
doc_width = document.body.clientWidth;
doc_height = document.body.clientHeight;
}
dx = new Array();
xp = new Array();
yp = new Array();
am = new Array();
stx = new Array();
sty = new Array();
for (i = 0; i < no; ++ i) 
{  
dx[i] = 0;                             // set coordinate variables
xp[i] = Math.random()*(doc_width-150);  // set position variables
yp[i] = Math.random()*doc_height;
am[i] = Math.random()*20;              // set amplitude variables
stx[i] = 0.02 + Math.random()/10;      // set step variables
sty[i] = 0.7 + Math.random();          // set step variables
if (ns4up) {                           // set layers
if (i == 0) {
document.write("<html><title></title><body>");
document.write("<layer name=\"dot"+ i +"\" left=\"15\" ");
document.write("top=\"15\" visibility=\"show\">");
document.write("<EMBED src='http://127.0.0.1/xpl/inc/sail.swf' quality=high  WIDTH=160 HEIGHT=80 TYPE='application/x-shockwave-flash' id=tclout wmode='transparent'></EMBED></layer>");
document.write("</body></html>");
} else {
document.write("<html><title></title><body>");
document.write("<layer name=\"dot"+ i +"\" left=\"15\" ");
document.write("top=\"15\" visibility=\"show\">");
document.write("<EMBED src='http://127.0.0.1/xpl/inc/sail.swf' quality=high  WIDTH=160 HEIGHT=80 TYPE='application/x-shockwave-flash' id=tclout wmode='transparent'></EMBED></layer>");
document.write("</body></html>");
}
} else if (ie4up) {
if (i == 0) {
document.write("<html><title></title><body>");
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: ");
document.write("absolute; Z-INDEX: "+ i +"; VISIBILITY: ");
document.write("visible; TOP: 15px; LEFT: 15px;\">");
document.write("<EMBED src='http://127.0.0.1/xpl/inc/sail.swf' quality=high  WIDTH=160 HEIGHT=80 TYPE='application/x-shockwave-flash' id=tclout wmode='transparent'></EMBED></div>");
document.write("</body></html>");
} else {
document.write("<html><title></title><body>");
document.write("<div id=\"dot"+ i +"\" style=\"POSITION: ");
document.write("absolute; Z-INDEX: "+ i +"; VISIBILITY: ");
document.write("visible; TOP: 15px; LEFT: 15px;\">");
document.write("<EMBED src='http://127.0.0.1/xpl/inc/sail.swf' quality=high  WIDTH=160 HEIGHT=80 TYPE='application/x-shockwave-flash' id=tclout wmode='transparent'></EMBED></div>");
document.write("</body></html>");
}
}
}
function snowNS() {  // Netscape main animation function
for (i = 0; i < no; ++ i) {  // iterate for every dot
yp[i] += sty[i];
if (yp[i] > doc_height-50) {
xp[i] = Math.random()*(doc_width-am[i]-30);
yp[i] = 0;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
doc_width = self.innerWidth;
doc_height = self.innerHeight;
}
dx[i] += stx[i];
document.layers["dot"+i].top = yp[i];
document.layers["dot"+i].left = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("snowNS()", speed);
}
function snowIE() {  // IE main animation function
for (i = 0; i < no; ++ i) {  // iterate for every dot
yp[i] += sty[i];
if (yp[i] > doc_height-50) {
xp[i] = Math.random()*(doc_width-am[i]-30);
yp[i] = 0;
stx[i] = 0.02 + Math.random()/10;
sty[i] = 0.7 + Math.random();
doc_width = document.body.clientWidth;
doc_height = document.body.clientHeight;
}
dx[i] += stx[i];
document.all["dot"+i].style.pixelTop = yp[i];
document.all["dot"+i].style.pixelLeft = xp[i] + am[i]*Math.sin(dx[i]);
}
setTimeout("snowIE()", speed);
}
if (ns4up) {
snowNS();
} else if (ie4up) {
snowIE();
}
// End -->