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
rs.open "select 數量 from 物品 where 物品名='礦石' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的現在還沒有礦石，所以你不可以進行練鐵操作！！');window.close();</script>"
	response.end
else
	if rs("數量")<2 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script Language=Javascript>alert('提示：您的現在還沒有鐵塊，所以你不可以進行練鐵操作！！');window.close();</script>"
		response.end
	end if
	Session("tiesl")=rs("數量")
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
<TITLE>采礦練鐵</TITLE>
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