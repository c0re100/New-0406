<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Session("iq")="" then
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>最後一關</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../chat/READONLY/Style.css" rel=stylesheet></head>

<body bgcolor="#000000" oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr> 
    <td width="630"> 
      <div align="left"><b><font color="#000000">最後一關了，寶藏的門前坐著一個年輕人  </font></b></div>
    </td>
    <td width="128" rowspan="2" valign="top"><img src="face3.GIF" width="128" height="290"></td>
  </tr>
  <tr> 
    <td width="630" valign="top"> 
      <p style="line-height: 200%; margin: 20"> <font color="#9966CC">年輕人：別問我是誰~！你隻要知道我是你的最後一個對手就可以了。前面的4關，你經過了棋力、智力、攻擊、賭術的考驗，現在有我來考驗你的內力，你能把我擊敗，才能用內力破開寶藏的大門，放馬過來吧，我不會留情的！</font> 
      <p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=info(0)%>:<a href="go9.asp">好~準備好輸給我吧！！！</a><br>
        </font> </p>
      <p style="line-height: 200%; margin: 20"> 
    </td>
  </tr>
  <tr valign="top"> 
    <td colspan="2"> </td>
  </tr>
</table>
</body>
</html>