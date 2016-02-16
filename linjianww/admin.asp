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
<%Response.Expires=0
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
if info(2)<10  then Response.Redirect "manerr.asp?id=100"
%>
<html>
<head>
<title>超級管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<%if info(0)=""  then
Response.Redirect "../error.asp?id=440"
elseif info(2)<10 then
Response.Redirect "../error.asp?id=439"
else%>
<frameset cols="110,652" rows="*"> 
  <frame src="menu.asp" noresize scrolling="YES" frameborder="NO" name="menu">
<frame src="txt.asp" noresize scrolling="AUTO" frameborder="NO" name="txt">
</frameset>
<noframes><body bgcolor="#FFFFFF">

</body></noframes>
</html>
<%end if%>
<html></html> 