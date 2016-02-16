<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
%>
<HTML><HEAD>
<!--include the remote scripting include files-->
<script language="Javascript" src="../_ScriptLibrary/rs.htm"></script>
<script language="Javascript">RSEnableRemoteScripting("../_ScriptLibrary");</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<title>江湖走鏢</title>
<body bgcolor="#000000" oncontextmenu="click2()">
<STYLE TYPE="text/css">
.item{}
.worm1 {
font-weight:bold;
font-size:12pt;
color:#ffffff;
background:#000000;
width:150px;
BORDER-BOTTOM: #ffffff solid 0px; 
BORDER-TOP: #ffffff solid 0px; 
BORDER-LEFT: #ffffff solid 0px; 
BORDER-RIGHT: #ffffff solid 0px; 
}
</style>
<SCRIPT LANGUAGE="JavaScript">
<%if session("jiutian_name")="" or session("jiutian_id")="" then%>
window.alert("您沒有登錄，或已經超時斷開了連接......");
top.location.href="../Index.asp";
</script>
<%
response.end
end if%>
var height = 20; 
var width = 20;
var speed = 100;
var missed = 1;
width += 2;
var a = 0;
var b = 0;
document.write("<table bgcolor=white align=center bordercolor=white");
document.write("align=center border=0 cellpadding=0 cellspacing=0><tr><td>");
for (b = 0; b < height+2; b++) {
document.write("<img src=end.gif width=0 height=0>");
for (a = 0; a < width- 2; a++) {
if ((b == 0) || (b == height+1)) {
document.write("<img src=end.gif width=0 height=0>");
}
else {
document.write("<img src=blank.gif>");
}
}
document.write("<img src=end.gif width=0 height=0><br>");
}
document.write("</td>");
document.write("<br></td></tr></table>");
document.write("<div align=center>");
document.write("<form name=info>");
document.write("<input type=button size=0 value=0 class=worm1></form></div>");
alert('摩卡提示:\n\n-摩卡老祖給了你一錠銀子說：出去闖闖吧.\n- 你必須把出現的寶物搶過來，換取銀子.\n- 每踫壁一次損傷10點生命力.\n- 上下左右鍵控制遊走路線.\n- 必須使用IE而不能使用多窗口瀏覽器.\n\n 按\"確定\" 開始....');
var points = 0;
var go = 1;
var di = 0;
var x = 0;
var y = 0;
var n = 0;
document.images[1].src = "blank.gif"; 
var blank = document.images[1].src;
var hw = (height * width);
var o = Math.floor(Math.random() * hw - 2);
do {
o = Math.floor(Math.random() * hw-2);
} while(document.images[o].src != blank);
var i = o;
var food = 0;
do {
food = Math.floor(Math.random() * hw-2);
} while (document.images[food].src != blank);
document.images[i].src = "worm.gif";
document.images[width-1].src="end.gif";
var end = document.images[width-1].src;
var file = document.images[i].src;
var length = 1;
var worm = new Array();
var k = 0;
var ie = document.all ? 1 : 0;
var enableScroll = ((navigator.appName == "Microsoft Internet Explorer") && (parseInt(navigator.appVersion) >3)) ;
var height = document.images[0].height;
var tScroll;
var d = 0;
function runTimer() {
if (d != 0) { n++; }
if (d == 1) { i--; }
if (d == 2) { i++; }
if (d == 3) { i += width; }
if (d == 4) { i -= width; }
if (document.images[i].src == end) {
speed -= 400; i = worm[n-1]; di = 1; die();
}
worm[n] = i;
if(i == food) {
length++; points += (1*length);
do {
food = Math.floor(Math.random() * hw-2);
} while (document.images[food].src != blank);
if (di == 0) {
document.info.elements[0].value = points;
}
}
if (n > length){
o = worm[n-length];
}
if ((document.images[i].src == file) && (n > 1)) {
speed -= 400; d = 0; di = 1; die();
}
if(di == 0) {
document.images[o].src = "blank.gif";
document.images[i].src = "worm.gif";
document.images[food].src = "q.gif";
tScroll = window.setTimeout("runTimer();", speed);
}
}
if (enableScroll){
if (ie) window.onload = runTimer;
if (ie) window.onunload = new Function("clearTimeout(tScroll)");
}
systm = "";
ver = navigator.appVersion;
len = ver.length;
for (iln = 0;iln < len; iln++) if (ver.charAt(iln) == "(") break;
systm = ver.charAt(iln+1).toUpperCase();
document.onkeydown = keyDown;
if (systm != "C") {
document.captureEvents(Event.KEYDOWN);
}
function keyDown(DnEvents) {
if (systm != "C") {
k = DnEvents.which;
} else {
k = window.event.keyCode;
}
if (k == 37) { d = 1; }
if (k == 39) { d = 2; }
if (k == 40) { d = 3; }
if (k == 38) { d = 4; }
}
function die() {
i = 0;
o = 0;
food = 0;
missed=10;
document.info.elements[0].value = points;
vsg(); 
}
var rsMath = RSGetASPObject("yb2.asp");
function tryremote()
{
try {
openx();
} 
catch(err)
{
Getobj();
openx();
}
}
function Getobj()
{
rsMath = RSGetASPObject("yb2.asp");
}
function vsg()
{
//找個理由不讓刷屏，哈哈
if(document.all("shoug").value=="領銀子")
{
document.all("shoug").value="你筋疲力盡帶了你的寶物回來........";
setTimeout('tryremote();',1);
}
}
function openx()
{
var ExeSql;
ExeSql=rsMath.addnl(points,missed);
window.alert(ExeSql.return_value);
window.location.href="../JIU/TWT.HTM";
}
function click2(){
window.location.href="../JIU/TWT.HTM";
window.alert("你慌裡慌張逃了出來，白忙了一場......");
} 
</script>
<div style="position:absolute;z-index:100;overflow: visible;">
<input type="button" onmouseup="vsg()" name="shoug" value="領銀子" style="background-color:#000000;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='#ffffff'">
</div>
</body>
</html>
