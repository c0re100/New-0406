<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("iq")="" then
response.redirect "index.asp"
end if
%>
<html>
<head>
<title>第四關</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

<LINK href="../../chat/READONLY/Style.css" rel=stylesheet></head>

<body oncontextmenu=self.event.returnValue=false background="images/bg.gif" >
<table width="600" border="0" cellspacing="0" cellpadding="0" align="center">
<tr>
<td width="617">
      <div align="left"><b><font color="#0000FF">剛走到第四關門前，一個獨眼龍就從樹後面轉了出來  </font></b></div>
</td>
<td width="136" rowspan="2" valign="top"><font color="#FFFFFF"><img src="face2.GIF" width="136" height="283"></font></td>
</tr>
<tr>
<td width="617" valign="top">
      <p style="line-height: 200%; margin: 20"> 司徒輸光：傻小子，想闖過我司徒輸光這一關可沒那麼容易。來來來，咱們來賭賭一把，要是你輸了，打哪兒來的就回哪兒去；要是你贏了，我就放你過去。怎麼樣，要不要試試？ 
      <p style="line-height: 200%; margin: 20"><font color="#FFFFFF"><%=info(0)%>:<a href="go7.asp">老子可是天下第一賭神，怕你什麼。</a><br>
</font>
</p>

</td>
</tr>
</table>
</body>
</html>