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
<!--#include file="homecheck.asp"-->
<!--#include file="checksession.asp"-->
<!--#include file="data.asp"-->
<html>
<%

if session("xiuxi")="1" then
conn.close
%>
<script>
window.alert("�@�Ѱ���𮧤@���I")
history.back()
</script>
<%
else
set ls=conn.execute("select * from �Τ� where �m�W='"&info(0)&"'")
house=ls("house")
select case house
case "house1":
conn.execute("update �Τ� set ��O=��O+30 where id="&info(9))
case "house2":
conn.execute("update �Τ� set ��O=��O+60 where id="&info(9))
case "house3":
conn.execute("update �Τ� set ��O=��O+90 where id="&info(9))
case "house4":
conn.execute("update �Τ� set ��O=��O+120 where id="&info(9))
case "house5":
conn.execute("update �Τ� set ��O=��O+150 where id="&info(9))
case "house6":
conn.execute("update �Τ� set ��O=��O+180 where id="&info(9))
end select
session("xiuxi")="1"
%>
<script>
window.alert("��O�w�g�K�[�F�I")
top.main.location.reload()
history.back()
</script>
<%
conn.close
end if
%>
</html>