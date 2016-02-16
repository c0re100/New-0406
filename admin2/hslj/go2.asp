<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")<>"asw" then
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>第二關</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet></head>

<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="595">
      <div align="left"><b><font color="#0000FF">第二關是一位PPMM耶~！！！</font></b></div>
</td>
<td width="165" rowspan="2" valign="top"><font color="#FFFFFF"><img src="jh0.jpg" width="153" height="172"></font></td>
</tr>
<tr>
<td width="595" valign="top">
      <p style="line-height: 200%; margin: 20"> PPMM：嗨，小帥哥，我父親是獨孤城的智囊。我來代我的父親把手第二關，不要怕呀，我這裡有十道問題，全部答對了就可以過去，要不要試試呀？</p>
</td>
</tr>
<tr valign="top">
<td colspan="2">
<p><font color=red><%=info(0)%>:</font><a href="answer.asp">好吧，我的外號可是叫做“問題天才”，把卷子給我。</a><br>
</p>
</td>
</tr>
</table>
</body>
</html>