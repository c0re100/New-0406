<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "請進入聊天室再進入股票！"
window.close()
</script>
<%Response.End
end if

%>

<html>

<head>
<title></title>
</head>

<body bgcolor="#0099FF">
<table width="100%" border="0">
  <tr>
    <td>
      <div align="center"><font size="+5">股票市場暫停交易！</font></div>
    </td>
  </tr>
</table>
<font color="#006633"> </font> 
</body>
</html>