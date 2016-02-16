<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><script>
if(!confirm("你有這個本事嗎？"))
window.location=window.close()
</script>
<%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>第一關</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet></head>

<body bgcolor="#0099CC" oncontextmenu=self.event.returnValue=false background="images/bg.gif" LEFTMARGIN="0" TOPMARGIN="0" >
<table width="562" border="0" cellspacing="0" cellpadding="0" align="center"> 
<tr> <td width="595"> <div align="left"><b>寶藏的第一關是獨孤求敗的摯友––棋痴張老道</b></div></td><td width="162" rowspan="2" valign="top"><img src="jh1.jpg" width="160" height="250"></td></tr> 
<tr> <td width="595" valign="top"> <p style="line-height: 200%; margin: 20"> <font color="#FF3366">張老道：</font>好小子，這裡是寶藏的第一關，我老人家這把年紀就不跟你動手了，咱們下盤棋，你贏了我就放你過去。<br><br> 
<font color=red><%=info(0)%></font>:<a href="go1.asp">好吧，我一定會下贏你的。</a><br> </p></td></tr> 
</table>
</body>
</html>