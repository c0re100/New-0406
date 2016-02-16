<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * from 物品 where 物品名='冰水' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
	Session("icesl")=0
else
	Session("icesl")=rs("數量")
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Session("icejl1")=0
Session("icets")=0
Session("cbsj")=true
%>
<HTML>
<HEAD>
<TITLE>極地采冰</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
</HEAD>
<FRAMESET border=0 cols=*,100 frameBorder=no frameSpacing=0 rows=*>
<FRAME name=fs scrolling=no src="ice.asp" frameSpacing=0 scrolling=no frameBorder=no>
<FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=50,50,0>
<FRAME border=0 frameBorder=no frameSpacing=0 name=ig scrolling=no src="icejl.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="icecz.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz1 scrolling=no src="about:blank">
</FRAMESET>
</FRAMESET>
</HTML>
