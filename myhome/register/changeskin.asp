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

<!--#include file="data2.asp"-->
<%
if request.form("skin")="" then
conn.close
%>
<script language="vbscript">
msgbox"�z������ܤ@���Y���I",0,"ĵ�i"
history.back
</script>
<%
response.end
end if
%>
<%
skin=request.form("skin")&"."&"gif"
'response.write skin
conn.execute"update userinfo set skin='"&skin&"' where user='"&info(0)&"'"
conn.close
%>
<script language="vbscript">
msgbox"�z�w�g���\��ܤF�Y���I",0,"ĵ�i"
history.back
</script>