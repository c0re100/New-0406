<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<HTML><HEAD>
<!--include the remote scripting include files-->
<script language="Javascript" src="../_ScriptLibrary/rs.htm"></script>
<script language="Javascript">RSEnableRemoteScripting("../_ScriptLibrary");</script>
<title>江湖打鏢</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="jh.css" rel=stylesheet>
<style>
.basic {position:absolute; visibility:visible; top:-3000px;}
</style>
</head>
<body bgcolor="#000000">
<SCRIPT LANGUAGE="JavaScript" >

var yoffset = 20; 
var isNS = (document.layers);
var _all = (isNS) ? 'outer.document.' : 'all.' ;
var _style = (isNS) ? '' : '.style' ;
var _visible = (isNS) ? 'show' : 'visible' ;
var py = 0, px = 0, xoffset, loaded = false, obj = new Object();
var target, field, rocket1, rocket2, rocket3, rocket4, rocket5, rocket6, bullet, explosion;
var oktoshoot = false, bulletlevel, speed = 10, missed = 0, fired = 0, hit = 0, accuracy = 0;
function init() {
target = eval('document.'+_all+'target'+_style);
field = eval('document.'+_all+'field'+_style);
rocket1 = eval('document.'+_all+'rocket1'+_style);
rocket2 = eval('document.'+_all+'rocket2'+_style);
rocket3 = eval('document.'+_all+'rocket3'+_style);
rocket4 = eval('document.'+_all+'rocket4'+_style);
rocket5 = eval('document.'+_all+'rocket5'+_style);
rocket6 = eval('document.'+_all+'rocket6'+_style);
bullet = eval('document.'+_all+'bullet'+_style);
explosion = eval('document.'+_all+'explosion'+_style);
resize();
if (isNS) {
target.moveTo(300,0);
}
else {
target.pixelLeft = 300;
target.pixelTop = 0;
}
target.visibility = _visible;
field.visibility = _visible;
alert('摩卡提示:\n\n- 你踏入漆黑的練功房，你必須擊毀飛過來的暗器.\n- 鼠標左鍵出擊，右鍵退出.\n- 每點擊一次消耗一點內力，擊中一次增加五點內力,丟失一個暗器丟失一點體力.\n- 按空格鍵休息查看結果.\n- 在狀態欄看丟失暗器的個數.\n- 練完點\"收功\"，否則內力增加不了嘍.\n\n 按\"確定\" 開始....');
loaded = true;
oktoshoot = true;
dispmisses();
animate();
}
function dispmisses() {
status = "Rockets missed: "+missed;
setTimeout('dispmisses()', 500);
}
function showstat() {
accuracy = (Math.round(hit/fired*100));
if (isNaN(accuracy)) {accuracy = 0};
alert('你體息了一會.....\n\n按\"確定\"繼續..\n\n 共出手: '+fired+'次 \n 擊毀暗器: '+hit+'個 \n 失敗: '+(fired-hit)+'次 \n 丟失暗器: '+missed+'個 \n 精確度: '+accuracy+'%');
}
function testandmoveX(rocketN) {
if (isNS) {
rocketN.clip.left = -rocketN.left-speed;
rocketN.clip.right = 500-rocketN.left;
if (rocketN.left+speed>500) {
rocketN.left = -60;
if (rocketN.top != -3000) { missed++; }
}
else {
rocketN.left+=speed;
}
if (rocketN.left == -60)rocketN.top = Math.floor(Math.random()*12)*20+30;
}
else {
rocketN.clip = "rect(0px, "+(500-speed-rocketN.pixelLeft)+"px, 15px, "+(-rocketN.pixelLeft-speed)+"px)"; 
if (rocketN.pixelLeft+speed > 500) {
if (rocketN.pixelTop != -3000) {missed++}
rocketN.pixelLeft = -60;
}else {
rocketN.pixelLeft+=speed;
}
if (rocketN.pixelLeft == -60)rocketN.pixelTop = Math.floor(Math.random()*12)*20+30;
}
}
function animate() {
for (i = 1; i <= 6; i++) {
testandmoveX(eval('rocket'+i));
}
setTimeout('animate()',100);
}
function shootrocket(x,y) {
fired++;
bulletlevel = 4;
bullettime(x,y);
}
function bullettime(x,y) {
if (bulletlevel > 0) {
if (isNS) {
bullet.clip.width = bulletlevel*4;
bullet.clip.height = bulletlevel*4;
bullet.moveTo(x-(bulletlevel*2),y-(bulletlevel*2));
if (bulletlevel == 1) {
for (r = 1;r <= 6;r++) {
tr = eval('rocket'+r);
if ( (bullet.top >= tr.top+2) && (bullet.top <= tr.top+9) && (bullet.left <= tr.left+57) && (bullet.left >= tr.left) ) {
explosion.moveTo(tr.left,tr.top);
tr.top = -3000;
hit++;
setTimeout('explosion.top = -3000', 400);
}
}
}
}
else {
bullet.width = bulletlevel*4;
bullet.height = bulletlevel*4;
bullet.pixelTop = y-(bulletlevel*2);
bullet.pixelLeft = x-(bulletlevel*2)
if (bulletlevel == 1) {
for (r = 1; r <= 6; r++) {
tr = eval('rocket'+r);
if ( (bullet.pixelTop >= tr.pixelTop+2) && (bullet.pixelTop <= tr.pixelTop+9) && (bullet.pixelLeft <= tr.pixelLeft+57) && (bullet.pixelLeft >= tr.pixelLeft) ) {
explosion.pixelLeft = tr.pixelLeft;
explosion.pixelTop = tr.pixelTop;
tr.pixelTop = -3000;
hit++;
setTimeout('explosion.pixelTop = -3000', 400);
}
}
}
}
bulletlevel--;
setTimeout('bullettime('+x+','+y+')',200);
}
else {
(isNS) ? bullet.top = -3000 : bullet.pixelLeft = -3000;
oktoshoot = true;
}
}
function resize() {
if (isNS) {
xoffset = (window.innerWidth-300)/2;
document.outer.moveTo((window.innerWidth-500)/2 , yoffset);
}
else {
xoffset = (document.body.clientWidth-300)/2
document.all.outer.style.pixelTop = yoffset;
document.all.outer.style.pixelLeft = (document.body.clientWidth-500)/2;
}
}
function movetarget(evnt) {
if (loaded) {
if (isNS) {
py = evnt.pageY-yoffset-59;
px = evnt.pageX-xoffset+59;
}
else {
py = event.clientY-yoffset-59;
px = event.clientX-xoffset+59;
}
if ((py >= 0) && (py <= 241)) { (isNS) ? target.top = py : target.pixelTop = py; };
if ((px >= 0) && (px <= 441)) { (isNS) ? target.left = px : target.pixelLeft = px };
return false;
}
}
function fire(mouse) {
if(event.button==2){click2();}
if (oktoshoot) {
oktoshoot = false;
(isNS) ? shootrocket(target.left+30,target.top+30) : shootrocket(target.pixelLeft+30,target.pixelTop+30) ; 
return false;
}
else {
}
}
function reloadNS() {
setTimeout('reloadNS_1()', 500);
}
function reloadNS_1() {
window.location.reload();
}
(isNS) ? window.onresize = reloadNS : window.onresize = resize;
function getkeypress(keypress) {
keyp = (isNS) ? keypress.which : window.event.keyCode;
if (keyp == 32)showstat();
return false;
}
if (isNS) window.captureEvents(Event.KEYPRESS);
if (isNS) document.captureEvents(Event.KEYPRESS);
window.onkeypress = getkeypress;
document.onkeypress = getkeypress;
if (isNS) document.captureEvents(Event.MOUSEMOVE);
document.onmousemove = movetarget;
if (isNS) document.captureEvents(Event.MOUSEDOWN);
if (isNS) window.captureEvents(Event.MOUSEDOWN);
document.onmousedown = fire;
window.onmousedown = fire;
window.onload = init;
var basic = 'visibility="visible" width="60" top="-3000" height="15"';
var txt = '';
if (isNS) {
txt += '<layer name="outer" visibility="visible" width="500">';
txt += '<layer name="target" visibility="hidden" z-index="12"><img src="target.gif"></layer>';
txt += '<layer name="field" visibility="hidden" bgcolor="black" z-index="1" width="500" height="300"></layer>';
txt += '<layer name="explosion" '+basic+'z-index="11"><img src="explosion.gif"></layer>';
txt += '<layer name="rocket1" '+basic+'left="-460" z-index="4"><img src="rocket.gif"></layer>';
txt += '<layer name="rocket2" '+basic+'left="-380" z-index="5"><img src="rocket.gif"></layer>';
txt += '<layer name="rocket3" '+basic+'left="-300" z-index="6"><img src="rocket.gif"></layer>';
txt += '<layer name="rocket4" '+basic+'left="-220" z-index="7"><img src="rocket.gif"></layer>';
txt += '<layer name="rocket5" '+basic+'left="-140" z-index="8"><img src="rocket.gif"></layer>';
txt += '<layer name="rocket6" '+basic+'left="-60" z-index="9"><img src="rocket.gif"></layer>';
txt += '<layer name="bullet" visibility="visible" width="10" top="-3000" height="10" bgcolor="white" z-index="10"></layer>';
txt += '</layer>';
}
else {
txt += '<div id="outer" style="position:absolute; width:600px; height:300px">';
txt += '<div id="field" style="position:absolute; top:0px; left:0px; visibility:hidden; width:500px; height:300px; z-index:2; background-color:black;"></div>';
txt += '<div id="target" style="position:absolute; visibility:hidden; z-index:12"><img src="target.gif"></div>';
txt += '<div id="explosion" style="z-index:11;" class="basic"><img src="explosion.gif"></div>';
txt += '<div id="rocket1" style="left:-460px; z-index:4" class="basic"><img src="rocket.gif"></div>';
txt += '<div id="rocket2" style="left:-380px; z-index:5" class="basic"><img src="rocket.gif"></div>';
txt += '<div id="rocket3" style="left:-300px; z-index:6" class="basic"><img src="rocket.gif"></div>';
txt += '<div id="rocket4" style="left:-220px; z-index:7" class="basic"><img src="rocket.gif"></div>';
txt += '<div id="rocket5" style="left:-140px; z-index:8" class="basic"><img src="rocket.gif"></div>';
txt += '<div id="rocket6" style="left:-60px; z-index:9" class="basic"><img src="rocket.gif"></div>';
txt += '<div id="bullet" style="z-index:10; background-color:white; width:10px; height:10px; font-size:1px;" class="basic"></div>';
txt += '</div>';
}
document.write(txt);
var rsMath = RSGetASPObject("fb2.asp");
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
rsMath = RSGetASPObject("fb2.asp");
}
function vsg()
{
//找個理由不讓刷屏，哈哈
if(document.all("shoug").value=="收功")
{
document.all("shoug").value="你微合雙眼，氣沉丹田，看上去正在收功........";
setTimeout('tryremote();',2);
}
}
function openx()
{
ExeSql2=rsMath.addnl(fired,hit,missed);
window.alert(ExeSql2.return_value);
window.location.href="javascript:window.close()";
}
function click2(){
window.location.href="javascript:window.close()";
window.alert("你慌裡慌張逃了出來，白忙了一場......");
} 
</script>
<br>
<br>
<div style="position:absolute;z-index:100;overflow: visible;">
<input type="button" onmouseup="vsg()" name="shoug" value="收功" style="background-color:#000000;color:#ffffff;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='#ffffff'">
</div>
</body>
