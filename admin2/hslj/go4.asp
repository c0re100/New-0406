<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")<>"opy" then
response.redirect "index.asp"
end if %>
<html>
<head>
<title>第三關</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet></head>

<body oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="599">
      <div align="left"><font color="#FFFFFF"><b><font color="#0000FF">你走到第三個關口前，突然跳出一個打扮的奇形怪狀小老頭~</font></b></font></div>
</td>
<td width="147" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face.gif" width="142" height="241"></font></td>
</tr>
<tr>
<td width="599" valign="top">
      <p style="line-height: 200%; margin: 20"> 打不死：臭小子，水準不錯嘛。已經到了第三關。我叫打不死，從三歲開始打架，到今天已經60多年了。想過我這關，就動動你的拳頭，讓我老頭活活筋骨。要是你的攻擊裡高過我，我就放你過去~！
      <p><font color="#FFFFFF"><%=info(0)%>:<a href="go5.asp">怕你什麼，我武功可是天下第一！！</a>
</font>
</p>
<p>
</p>
<p> </p>
</td>
</tr>
</table>
</body>
</html>
