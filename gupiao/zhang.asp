<!--#include file="jhb.asp"-->
<%
if info(0)="" then
%>
<script language="vbscript">
  alert "����i�J�A�A�٨S���n��"
  location.href = "../index.asp"
</script>
<%
else
username=info(0)
if username<>"root" then%>
<script language="vbscript">
  alert "�A�S���ާ@"
  location.href = "index.asp"
</script>
<%
else
sid=Request.Form ("sid")
sql="update �Ѳ� set ��e����=��e����*1.1 where sid="&sid
conn.execute sql
conn.close
Response.Redirect "exchange.asp"
set conn=nothing
end if

