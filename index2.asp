<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
randomize timer
s=1+int(rnd*10000)%>
<%randomize timer
mysound=int(rnd*8)+1
Response.Write "<bgsound src=mid/midi" & mysound & ".mid loop=-1 volume=100>"
%>
<!--#include file="config.asp"-->
<script language="JavaScript" type="text/JavaScript">
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
</script>
<SCRIPT LANGUAGE="JavaScript">
function check(got) {
if (got.name.value=="") {
alert("請輸入用戶名 !!");
return false;
}
if (got.pass.value=="") {
alert("請輸入密碼 !!");
return false;
}
return true;
}
function openwin08(){
window.open('','ewin',"height=400,width=575,screenX=250,screenY=150,top=200,left=250,resizable=no,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");
}
function openwin07(){
window.open('','fwin',"height=360,width=570,screenX=250,screenY=150,top=100,left=250,resizable=no,scrollbars=yes,status=no,toolbar=no,menubar=no,location=no");
}
</script>
<HTML><HEAD><TITLE>新0406江湖</TITLE>
<META http-equiv=Content-Type content="text/html; charset=big5">
<BASE onmouseover="window.status=' 歡迎大家黎到本江湖,一注冊即可做會員仲唔叫多D~`friend玩!!!!!';return true">
<LINK href="001/dob.css" type=text/css rel=stylesheet>
<bgsound src="LEE%20CHUN%20WING/江湖程式/0406/0406/新0406/MID/Midi1.mid" loop="-1">

<BODY leftMargin=0 background=001/bg.gif topMargin=0 marginheight="0" marginwidth="0" scroll="no">
<DIV align=center style="width: 916; height: 516">
<img border="0" src="logo.gif"><br>
<TABLE cellSpacing=0 cellPadding=0 width=600 border=0>
<TR vAlign=top>
<TD width=6 background=001/bian2.gif></TD></TR></TABLE>
<TABLE cellSpacing=0 cellPadding=0 width=712 border=0>
<TR vAlign=top>
<TD width=6 background=001/bian.gif height=106>
<!-------鎖右鍵_開始-------->
<script language="JavaScript"> 
<!--   
if (navigator.appName.indexOf("Internet Explorer") != -1)   
document.onmousedown = noSourceExplorer;   
function noSourceExplorer(){if (event.button == 2 | event.button == 3)   
{   
alert("多謝你黎到new~~0406^^");}}   
//-->    
          </script>
<!-------鎖右鍵_結束--------></TD>
<TD width=756 background=001/bg33.jpg>
<TABLE cellSpacing=0 cellPadding=8 width="100%" border=0 height="464">
<TR>
<TD width="43%" vAlign=top height="448"><div align="center">
<table width="95%"border="0">
<tr>
<td><div align="center">
<table width="98%" border="0" cellpadding="2" cellspacing="2" bordercolor="#006666"  bgcolor="#FFFFFF"  style="BORDER-COLLAPSE: collapse ;border: 1px dashed  #000000">
<form method=POST action=check.asp  onSubmit="return(check(this))">
<input type=hidden name=reg1 value="<%=s%>">
<input type=hidden name=reg value="<%=s%>">
<tr bgcolor=FFFFCC>
<td width="27%" height="18" bgcolor="#F0F8FF"> <img border="0" src="001/a1.gif"><img border="0" src="001/c3.gif"></td>
<td width="73%" bgcolor="#F0F8FF"><script language=JScript src="online.asp"></script><select name=select>

");<%
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
onlinenow=0
for i=0 to chatroomnum	
	online=split(trim(Application("ljjh_useronlinename"&i))," ")
	onlinenum=ubound(online)+1
	chatroomxx=split(ljjh_roominfo(i),"|")
	onlinenow=onlinenow+onlinenum
next 
%>
document.write("<option selected style='color:#000000; '>聊天室共<%=onlinenow%>人在戰鬥</option>

");<%
for i=0 to chatroomnum	
	online=split(trim(Application("ljjh_useronlinename"&i))," ")
	onlinenum=ubound(online)+1
	chatroomxx=split(ljjh_roominfo(i),"|")
%>
document.write("<option style='font-size:9pt;color:#FFFFFF;background-color:42774F'><%=chatroomxx(0)%><%=onlinenum%>人聊天</option>

");<%	if onlinenum<>0 then
useronlinename=replace(trim(Application("ljjh_useronlinename"&i))," ","</option><option style='font-size:9pt;color:#000000;'>")
%>
document.write("<option style='color:#000000; '><%=useronlinename%></option>
");<%end if
next
%>
document.write("</select></td>
</tr>
<tr bgcolor=FFFFCC>
<td bgcolor="#F0F8FF"><img border="0" src="001/a3.gif"><img border="0" src="001/c4.gif"></td>
<td bgcolor="#F0F8FF"><input type="text" name="name" maxlength="12"  size="20" class='tinyBorder' style='width: 110;'></td>
</tr>
<tr bgcolor=FFFFCC>
<td height="25" bgcolor="#F0F8FF"> <img border="0" src="001/a5.gif"><img border="0" src="001/c5.gif"></td>
<td bgcolor="#F0F8FF"><input type="password" name="pass" maxlength="12"  size="20" class='tinyBorder' style='width: 110;'></td>
</tr>
<tr bgcolor=FFFFCC>
<td colspan="2" bgcolor="#F0F8FF"><span style="font-size: 9pt">
<marquee width="262">快進入本江湖玩吧，本網有免費１－4級會員，又有大量插件，如：購買數值、領取系統、釣魚插件、挖寶插件等，還在呆？快進入本江湖玩吧！</marquee></span></td></tr>
<tr bgcolor=FFFFCC>
<td colspan="2" bgcolor="#F0F8FF"><div align="center">
<input type="submit" name="Submit" value="登錄江湖" class='tinyBorder' style='width: 60;'>
<input type="reset" name="Submit2" value="重新填寫" class='tinyBorder' style='width: 60;'>
</div></td></tr></form>
</table>
</div></td>
</tr>
</table>
<br>
<table cellSpacing="0" cellPadding="0" width="80%" border="0">
  <tr>
    <td align="middle" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    <a href="jh.asp" target="_blank" style="color: #FFFFFF; text-decoration: none">
    <img height="39" src="aa.gif" width="148" border="0"></a></td>
  </tr>
  <tr>
    <td align="middle" height="8" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    </td>
  </tr>
  <tr>
    <td align="middle" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    <a href="yamen/modify.asp" target="_blank" style="color: #FFFFFF; text-decoration: none">
    <img height="39" src="cc.gif" width="148" border="0"></a><a href="yamen/disp.asp" target="_blank" style="color: #FFFFFF; text-decoration: none"><img height="39" src="dd.gif" width="148" border="0"></a></td>
  </tr>
  <tr>
    <td align="middle" height="8" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    </td>
  </tr>
  <tr>
    <td align="middle" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    <a href="yamen/read.htm" target="_blank" style="color: #FFFFFF; text-decoration: none">
    <img height="61" src="bb.gif" width="148" border="0"></a><a href="yamen/close.asp" target="_blank" style="color: #FFFFFF; text-decoration: none"><img height="61" src="ee.gif" width="148" border="0"></a></td>
  </tr>
  <tr>
    <td align="middle" height="8" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    </td>
  </tr>
  <tr>
    <td align="middle" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    　</td>
  </tr>
  <tr>
    <td align="middle" height="8" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    </td>
  </tr>
  <tr>
    <td align="middle" style="font-family: 宋體; font-size: 12px; line-height: 140%">
    　</td>
  </tr>
</table>
<p>　</div></TD>
<TD width="57%" height=448 vAlign=top> 
<table width="392" cellpadding="4" bgcolor=42774F cellspacing-="4" height="203">
<tr>
<td height="193" style="border:2px dashed white;" bgcolor="#F0F8FF" width="376">
<div align="center">
<table width="370" cellpadding="4" cellspacing="1" bgcolor=000000 height="362">
<tr bgcolor="#FFFFFF">
<td width="366" height="353" colspan="2" valign="top" background="ff.jpg">
　      
<iframe name="I1" width="360" height="330" src="../new_news/new.asp" border="0" frameborder="0" target="_self" align="center">
您的瀏覽器不支援內置框架或目前的設定為不顯示內置框架。</iframe>         
</td>
</tr>
</table>
</div></td>
</tr>
</table>
</TD>
</TR>
</TABLE></TD>
</TR></TABLE>
</DIV>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </p>
 <script language="JavaScript">
var step=12;
var speed=10;
var message="歡迎光臨~新0406江湖";
var x,y;
var flag=0;
message=message.split("");
var xpos=new Array();
for (i=0;i<=message.length-1;i++) {
xpos[i]=-50;
}
var ypos=new Array();
for (i=0;i<=message.length-1;i++) {
ypos[i]=-50;
}
function handlerMM(e) {
x = (document.layers) ? e.pageX : document.body.scrollLeft+event.clientX+20;
y = (document.layers) ? e.pageY : document.body.scrollTop+event.clientY;
flag=1;
}
function makesnake() {
if (flag==1 && document.all) {
for (i=message.length-1; i>=1; i--) {
xpos[i]=xpos[i-1]+step;
ypos[i]=ypos[i-1];
}
xpos[0]=x+step;
ypos[0]=y;
for (i=0; i<=message.length-1; i++) {
var thisspan = eval("span"+(i)+".style");
thisspan.posLeft=xpos[i];
thisspan.posTop=ypos[i];
thisspan.color=Math.random() * 255 * 255 * 255 + Math.random() * 255 * 255 + Math.random() * 255;
}
}
else if (flag==1 && document.layers) {
for (i=message.length-1; i>=1; i--) {
xpos[i]=xpos[i-1]+step;
ypos[i]=ypos[i-1];
}
xpos[0]=x+step;
ypos[0]=y;
for (i=0; i<message.length-1; i++) {
var thisspan = eval("document.span"+i);
thisspan.left=xpos[i];
thisspan.top=ypos[i];
thisspan.color=Math.random() * 255 * 255 * 255 + Math.random() * 255 * 255 + Math.random() * 255;
}
}
}
for (i=0;i<=message.length-1;i++) {
document.write("<span id='span"+i+"' class='spanstyle'>");
document.write(message[i]);
document.write("</span>");
}
if (document.layers) {
document.captureEvents(Event.MOUSEMOVE);
}
document.onmousemove = handlerMM;
function pageonload() {
makesnake();
window.setTimeout("pageonload();", speed);
}
window.onload=pageonload;
</script>
    <style>
.spanstyle{font-size: 9pt; position: absolute; top: -50px; visibility: visible}
</style></BODY></HTML>
<!-- #include file="global.asp" -->