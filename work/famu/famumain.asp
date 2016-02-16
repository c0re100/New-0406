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
rs.open "select 職業 from 用戶 where id="&info(9),conn
if trim(rs("職業"))="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的職業有錯，所以您不能在這裡進行伐木！！');window.close();</script>"
	response.end
end if
rs.close
rs.open "select 數量 from 物品 where 物品名='木頭' and 擁有者="& info(9),conn
If Rs.Bof OR Rs.Eof then
	Session("minesl")=0
else
	Session("minesl")=rs("數量")
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Session("minets")=0
Session("minesj")=true
%>
<HTML>
<HEAD>
<TITLE>後山伐木</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<META content="MSHTML 5.00.2920.0" name=GENERATOR>
</HEAD>
<FRAMESET border=0 cols=* rows=*,80 frameBorder=no frameSpacing=0>
<FRAME name=fs scrolling=no src="famu.asp" frameSpacing=0 scrolling=no frameBorder=no>
<FRAMESET border=0 cols=* frameBorder=no frameSpacing=0 rows=50,50,0>
<FRAME border=0 frameBorder=no frameSpacing=0 name=ig scrolling=no src="famujl.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz scrolling=no src="famucz.asp">
<FRAME border=0 frameBorder=no frameSpacing=0 name=cz1 scrolling=no src="about:blank">
</FRAMESET>
</FRAMESET>
</HTML>