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
set rs=conn.execute("select * from userinfo where user='"&info(0)&"'")%>
�a�����A�G<%
if trim(rs("homeopen"))<>"0" then
%>
���}<%else%>�W��
<%end if%>
</tD>
<td width="100%" height="20" bgcolor="#FFFBEF" colspan="6">
<%
if trim(rs("homeopen"))<>"0" then
%>
<a href="homeopen_post.asp?action=close">��W�a��</a>
<%else%>
<a href="homeopen_post.asp?action=open">���}�a��</a>
<%end if%>
<%
set rs=nothing
conn.close
%>

