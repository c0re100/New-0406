<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if
if  info(2)<10 then
	Response.Redirect "../error.asp?id=439"
else%>
<!--#include file="jhb.asp"-->
<%
sid=Request.Form ("sid")
sql="update �Ѳ� set ��e����=��e����*0.9 where sid="&sid
conn.execute sql
conn.close
Response.Redirect "exchange.asp"
set conn=nothing
end if
%>
