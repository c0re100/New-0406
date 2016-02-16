<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>江湖美容院</title>
<meta content="text/html; charset=big5" http-equiv="Content-Type">
<meta content="zh-cn" http-equiv="Content-Language">
<meta content="Microsoft FrontPage 4.0" name="GENERATOR">
<meta content="FrontPage.Editor.Document" name="ProgId">
<style type="text/css">body         { font-family: 新細明體; font-size: 9pt }
td           { font-family: 新細明體; font-size: 9pt }
a:link       { text-decoration: underline }
a:hover      { text-decoration: none }
.words       { line-height: 15pt }
</style>
</head>

<body topmargin="0" leftmargin="0" BGCOLOR="#660000">
<div align="center"> <table width="514" border="0" cellspacing="0" cellpadding="0" height="68"> 
<tr bgcolor="#000000" > <td BGCOLOR="#996666" HEIGHT="30"><DIV ALIGN="CENTER"><font color="#000000"> <font size="2" color="#FFFFFF"><%=info(0)%>這兒是江湖美容院，這裡的小姐會給你增加魅力的哦。</font></font></DIV></td></tr> 
</table></div><table border="0" cellpadding="0" cellspacing="0" align="center"
width="514"> <tbody> <tr> <td align="middle" height="139"><a href="#"
title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-4s.jpg" width="120" height="120"></a></td><td align="middle" height="139"><a href="#"
title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-2s.jpg" width="120" height="120"></a></td><td align="middle" height="139"><a href="#"
title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-3s.jpg" width="120" height="120"></a></td><td align="middle" height="139"><a href="#"
title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-1s.jpg" width="120" height="120"></a></td><td align="middle" rowspan="2">&nbsp;</td></tr> 
<tr> <td align="middle" valign="top"><a
href="#" title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-5s.jpg" width="120" height="120"></a></td><td align="middle" valign="top"><a
href="#" title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-6s.jpg" width="120" height="120"></a></td><td align="middle" valign="top"><a
href="#" title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-8s.jpg" width="120" height="120"></a></td><td align="middle" valign="top"><a
href="#" title="我要美容"
onClick="window.open('meiyong1.asp?my=<%=info(0)%>','youyi','scrollbars=yes,resizable=yes,width=180,height=50')"><img
border="0" src="3-7s.jpg" width="120" height="120"></a></td></tr> </tbody> </table><div align="center"> 
</div>
</body>

</html>
<html></html>