<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)>=6  then %>
<!--#include file="const.inc"-->
<%
Response.Buffer=true
PK="<font color=9FDF70 >PK（比武）時間到，開始PK，請新手和不願意打架的開啟練功保護，以防誤殺！！！</font>"

%>
<html>
<head>
<style>
.P{ font-size: 40px}
.l{line-height:120%}
body{font-size:120pt}

</style>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>開始PK </title>

</head>

<body bgcolor=#DEB887 >
<center>
<tr><td>
<%
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)="大家"
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="對"
sd(199)="<br><div align='center'><img src='img/bad.gif' border='0'>"& PK &"</font><img src='img/bad.gif' border='0'></div><br>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%><script language="vbscript">
  alert "執行PK開始操作成功！"
window.close()
</script><%
%>
</body>
</html>
<%end if%>