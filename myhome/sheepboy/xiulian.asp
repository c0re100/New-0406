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

<!--#include file="data2.asp"-->
<%
if info(0)="" then
response.redirect"warning.htm"
else
sheepname=request.form("sheepname")
rs.open"select user from sheep where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')",conn,1,1
%>
<!--#include file="data.asp"-->
<%
conn.execute "update 用戶 set 銀兩=銀兩-100000 where 姓名='"&info(0)&"'"
conn.close%>
<!--#include file="data1.asp"-->
<%
conn.execute"update 孩子技能 set 內力=1000 where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
conn.close
%>
<script language="vbscript">
msgbox"陪伴完畢！",0,"FLUSH"
history.back
</script>
<%
end if
%>