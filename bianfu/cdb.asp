<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html><head>
<meta http-equiv="refresh" content="10">
<style></style>
<script language="JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
// -->
</script>

<style type="text/css">

#superball {
position:absolute;
left:0;
top:0;
visibility:hide;
visibility:hidden;
width:40;
height:40;
}

</style>



<STYLE TYPE="text/css">
<!--

BODY{
overflow:scroll;
overflow-x:hidden;
}

.s1
{
  position  : absolute;
  font-size : 12pt;
  color     : blue;
  visibility: hidden;
}

.s2
{
  position  : absolute;
  font-size : 20pt;
  color     : red;
	visibility : hidden;
}

.s3
{
  position  : absolute;
  font-size : 16pt;
  color     : gold;
	visibility : hidden;
}

.s4
{
  position  : absolute;
  font-size : 14pt;
  color     : lime;
	visibility : hidden;
}
-->
</STYLE>


</head>
<BODY oncontextmenu=self.event.returnValue=false background=ying2.jpg>
<%
if Session("diaoyu")=true then
	Session("diaoyu")=now()
end if 
abc=DateDiff("s",Session("diaoyu"),now())
if abc>=115 then
%>
<div align="center">
<script language=vbscript>
location.href = "pao2.asp"
</script>
<%
response.end
end if
if abc<=90 then
%>
<br>
    <font color="#FFFFFF" size="2">【靈劍江湖】派江湖保鏢幫你去拿槍了,<br>他已經去了: <%=abc%>秒鐘<br>慢慢等吧,靈劍江湖保鏢會把槍拿來的!</font>
<%
else
%>
<IMG SRC="qiang.gif" border="0" width="80" height="80">     
    <font color="#0000FF"><br></font> 
    <font color="#FFFFFF" size="2">哈哈,靈劍江湖保鏢幫你把槍拿來了,快打呀!</font>
<%
end if
%>
 
<DIV ID="div1" CLASS="s1">*</DIV> 
<DIV ID="div2" CLASS="s2">*</DIV> 
<DIV ID="div3" CLASS="s3">*</DIV> 
<DIV ID="div4" CLASS="s4">*</DIV> 
 
 
<script language="javascript" type="text/javascript" class="s5"> 
 
var nav = (document.layers); 
var tmr = null; 
var spd = 50; 
var x = 0; 
var x_offset = 5; 
var y = 0; 
var y_offset = 15; 
 
if(nav) document.captureEvents(Event.MOUSEMOVE); 
document.onmousemove = get_mouse;  
 
function get_mouse(e) 
{     
  x = (nav) ? e.pageX : event.clientX+document.body.scrollLeft; 
  y = (nav) ? e.pageY : event.clientY+document.body.scrollTop; 
  x += x_offset; 
  y += y_offset; 
  beam(1);      
} 
 
function beam(n) 
{ 
  if(n<5) 
  { 
    if(nav) 
    {  
      eval("document.div"+n+".top="+y); 
      eval("document.div"+n+".left="+x); 
      eval("document.div"+n+".visibility='visible'"); 
    }   
    else 
    { 
      eval("div"+n+".style.top="+y); 
      eval("div"+n+".style.left="+x); 
      eval("div"+n+".style.visibility='visible'"); 
    } 
    n++; 
    tmr=setTimeout("beam("+n+")",spd); 
  } 
  else 
  { 
     clearTimeout(tmr); 
     fade(4); 
  }    
}  
 
function fade(n) 
{ 
  if(n>0)  
  { 
    if(nav)eval("document.div"+n+".visibility='hidden'"); 
    else eval("div"+n+".style.visibility='hidden'");  
    n--; 
    tmr=setTimeout("fade("+n+")",spd); 
  } 
  else clearTimeout(tmr); 
}  
 
// --> 
</script> 
 
<p>  
<script language="JavaScript">  
<!--//  
  
//設置下面一些參數，小球移動速度1-50，數值大速度快；  
var ballWidth = 100;  
var ballHeight = 100;  
var BallSpeed = 30;  
  
var maxBallSpeed = 200;  
var xMax;  
var yMax;  
var xPos = 80;  
var yPos = 50;  
var xDir = 'right';  //水平方向向右移動  
var yDir = 'down'; //垂直方向向下移動  
var superballRunning = true;  
var tempBallSpeed;  
var currentBallSrc;  
var newXDir;  
var newYDir;  
  
function initializeBall() {  
   if (document.all) {  
      xMax = document.body.clientWidth  
      yMax = document.body.clientHeight  
      document.all("superball").style.visibility = "visible";  
      }  
   else if (document.layers) {  
      xMax = window.innerWidth;  
      yMax = window.innerHeight;  
      document.layers["superball"].visibility = "show";  
      }  
   setTimeout('moveBall()',400);  
   }  
  
function moveBall() {  
   if (superballRunning == true) {  
      calculatePosition();  
      if (document.all) {  
         document.all("superball").style.left = xPos + document.body.scrollLeft;  
         document.all("superball").style.top = yPos + document.body.scrollTop;  
         }  
      else if (document.layers) {  
         document.layers["superball"].left = xPos + pageXOffset;  
         document.layers["superball"].top = yPos + pageYOffset;  
         }  
      setTimeout('moveBall()',30);  
      }  
   }  
  
function calculatePosition() {  
   if (xDir == "right") {  
      if (xPos > (xMax - ballWidth - BallSpeed)) {   
         xDir = "left";  
         }  
      }  
   else if (xDir == "left") {  
      if (xPos < (0 + BallSpeed)) {  
         xDir = "right";  
         }  
      }  
   if (yDir == "down") {  
      if (yPos > (yMax - ballHeight - BallSpeed)) {  
         yDir = "up";  
         }  
      }  
   else if (yDir == "up") {  
      if (yPos < (0 + BallSpeed)) {  
         yDir = "down";  
         }  
      }  
   if (xDir == "right") {  
      xPos = xPos + BallSpeed;  
      }  
   else if (xDir == "left") {  
      xPos = xPos - BallSpeed;  
      }  
   else {  
      xPos = xPos;  
      }  
   if (yDir == "down") {  
      yPos = yPos + BallSpeed;  
      }  
   else if (yDir == "up") {  
      yPos = yPos - BallSpeed;  
      }  
   else {  
      yPos = yPos;  
      }  
   }  
  
if (document.all||document.layers)  
window.onload = initializeBall;  
window.onresize = new Function("window.location.reload()");  
  
// -->  
</script>  
  
<p>  
  
<span id="superball"><a href="cdbok.asp"><img name="superballImage" src="ying.gif" height=100 width=100 border="0"></a></span>                 
</body>  
</html> 