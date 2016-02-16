<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
%>
<SCRIPT language="JavaScript">
<!--
bname=navigator.appName;
bversion=parseInt(navigator.appVersion)
if (bname=="Netscape")
brows=true
else
brows=false
var x=0;
var link=new Array();

function dspl(msg,bgcolor,dtop,delft){
this.msg=msg;
this.bgcolor=bgcolor;
this.dtop=dtop;
this.dleft=delft;
}


link[0]=new dspl('<div align="center"><font color="#FF6600">官 府 網 管 資 料</div><br><div align="center" > 姓 名&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;等 級&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;狀 態</div></font><%
rs.open "SELECT * FROM 用戶 where 門派='官府' order by grade DESC",conn
do while not rs.eof and not rs.bof
%><div align="center"> <%=rs("姓名")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("grade")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%=rs("狀態")%></div><%rs.movenext
loop%>','bisque',90,230)
<%

rs.close
set rs=nothing
conn.close
set conn=nothing
%> 
link[1]=new dspl('<div align="center"><font color="#FF6600" size="2">管理員權利！</font></div><li><font color="#000000" size="2">管理員隻管理刷屏、罵人，及違反中國有關互聯網規定的對像！</font></li><li><font color="#000000" size="2">官府身份分為6--10級，按照身份的高低擁有不同的權力！</font></li><li><font color="#000000" size="2">警告與踢人為基本權力！</font></li><li><font color="#000000" size="2">6級以上身份可以逮捕！</font></li><li><font color="#000000" size="2">7級以上可以抓人坐牢，10分鐘不可進江湖(踢人)！</font></li><li><font color="#000000" size="2">8級以上可以發錢（暫時封閉這個功能，怕亂）！</font></li><li><font color="#000000" size="2">9級以上可以使用特別管理系統！！</font></li><li><font color="#000000" size="2">10級為超級管理員！</font></li><li><font color="#000000" size="2">發現管理員有亂用權力的，應當及時舉報投訴，站長將在查清事實後進行嚴肅的處理！！</font></li>','bisque',90,230)

// Do not edit anything else in the script !!!!

function don(x){
if ((bname=="Netscape" && bversion>=4) || (bname=="Microsoft Internet Explorer" && bversion>=4)){
if (brows){
with(link[x]){
document.layers['linkex'].bgColor=bgcolor;
document.layers['linkex'].document.writeln(msg);
document.layers['linkex'].document.close();
document.layers['linkex'].top=dtop;
document.layers['linkex'].left=dleft;
}
document.layers['linkex'].visibility="show";
}
else{
with(link[x]){
linkex.innerHTML=msg;
linkex.style.top=dtop;
linkex.style.left=dleft;
linkex.style.background=bgcolor;
}
linkex.style.visibility="visible";
}
}
}

function doff(){
if ((bname=="Netscape" && bversion>=4) || (bname=="Microsoft Internet Explorer" && bversion>=4)){
if (brows)
document.layers['linkex'].visibility="hide";
else
linkex.style.visibility="hidden";
}
}

//-->


</SCRIPT>
<html>
<head>
<title>官府衙門</title>
<link rel="stylesheet" href="../css.css">
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
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor="#FFFFFF" background="../001/tanpopo7.jpg" leftmargin="0" topmargin="0">
<div id="linkex" style="position: absolute; visibility: hidden; width=80%; left: 152px; top: 115px; width: 267px; height: 264px"> 
</div>
<p align="center"><br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <img src="../001/kadomatu3.gif" width="270" height="212" usemap="#Map" border="0"></p>
<p align="center"><font color="#000000" size="-1"><br>
  </font> 
  <map name="Map"> 
    <area shape="rect" coords="479,353,558,396" a href ="" onClick="window.close()" alt="退出衙門" title="退出衙門"> 
    <area shape="rect" coords="104,149,147,189" href="#" onmouseover="don(1)" onmouseout="doff()" target="_top" >
    <area shape="rect" coords="157,128,190,187" href="#" onmouseover="don(0)" onmouseout="doff()" target="_top" >
  </map>
</p>
<script language=javascript> 
     function Click(){
     alert('靈劍江湖歡迎您!!!');
     window.event.returnValue=false;
     }
     document.oncontextmenu=Click;
     </script>
</body>

</html>

 