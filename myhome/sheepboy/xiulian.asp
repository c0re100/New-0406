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
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-100000 where �m�W='"&info(0)&"'"
conn.close%>
<!--#include file="data1.asp"-->
<%
conn.execute"update �Ĥl�ޯ� set ���O=1000 where sheepname='"&sheepname&"' and (user='"&info(0)&"' or sheep002='"&info(0)&"')"
conn.close
%>
<script language="vbscript">
msgbox"���񧹲��I",0,"FLUSH"
history.back
</script>
<%
end if
%>