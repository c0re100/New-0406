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
rs.open "select �Ȩ� from �Τ� where ID="&INFO(9),conn
money=rs("�Ȩ�")+500000
conn.execute("update �Τ� set �Ȩ�='"&money&"'  where ID="&INFO(9))
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
�A�w�g���\�o��50�U���I
<%else%>
�A�S������!
<%rs.close
set rs=nothing	
conn.close
set conn=nothing
end if%>