<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>���D�^��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style.css" rel=stylesheet>
</head>

<body background="images/bg.gif">
<table width="595" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td><%dim r,w 
r=0
w=0
for i=1 to 10 
ann=request.form("ann"&I&"")
an=request.form("an"&I&"")
response.write "��"&I&"�D"
if ann=an then 
response.write "���T<br>"
r=r+1
else
response.write "���I���׬O"&ann&"<br>"
w=w+1
end if
next
if r>7 then
Session("iq")="opy"
response.redirect "go4.asp"
else
response.write "���A¸��E�Y�r!�A�S�L���A�I<a href='index.asp'>�o��</a>�^�h�a~"
end if
%></td>
  </tr>
</table>
</body>
</html>