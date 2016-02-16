<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data1.asp"-->
<%set rs=conn.execute("select user from sheep where user='"&info(0)&"'")
if not rs.eof then
rs.close
conn.execute("delete from sheep where user='"&info(0)&"'")
conn.close
%>
<!--#include file="data.asp"--><%
rs.open "select 銀兩 from 用戶 where ID="&INFO(9),conn
money=rs("銀兩")+500000
conn.execute("update 用戶 set 銀兩='"&money&"'  where ID="&INFO(9))
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
你已經成功得到50萬錢！
<%else%>
你沒有收獲!
<%rs.close
set rs=nothing	
conn.close
set conn=nothing
end if%>