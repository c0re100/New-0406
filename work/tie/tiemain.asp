<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ƶq from ���~ where ���~�W='�q��' and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('���ܡG�z���{�b�٨S���q�ۡA�ҥH�A���i�H�i��m�K�ާ@�I�I');window.close();</script>"
	response.end
else
	if rs("�ƶq")<2 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('���ܡG�z���{�b�٨S���K���A�ҥH�A���i�H�i��m�K�ާ@�I�I');window.close();</script>"
		response.end
	end if
	Session("tiesl")=rs("�ƶq")
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Session("tiets")=0
Session("tiesj")=true
%>
<HTML>
<HEAD>
<TITLE>���q�m�K</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
</HEAD>
<FRAMESET border=0 cols=* rows=*,80 frameBorder=no frameSpacing=0>
<FRAME name=fs scrolling=no src="tie.asp" frameSpacing=0 scrolling=no frameBorder=no>
<FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=50,50,0>
<FRAME border=0 frameBorder=no frameSpacing=0 name=ig scrolling=no src="tiejl.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="tiecz.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz1 scrolling=no src="about:blank">
</FRAMESET>
</FRAMESET>
</HTML>