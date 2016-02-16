<!--#include file="jhb.asp"-->
<%
if info(0)="" then
%>
<script language="vbscript">
  alert "不能進入，你還沒有登錄"
  location.href = "../index.asp"
</script>
<%
else
username=info(0)
if username<>"root" then%>
<script language="vbscript">
  alert "你沒資格操作"
  location.href = "index.asp"
</script>
<%
else
sid=Request.Form ("sid")
sql="update 股票 set 當前價格=當前價格*1.1 where sid="&sid
conn.execute sql
conn.close
Response.Redirect "exchange.asp"
set conn=nothing
end if

