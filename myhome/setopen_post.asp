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
<!--#include file="data1.asp"-->

<%
rs.open "select * from openlist",conn,1,3
if not rs.bof then
conn.execute "delete from openlist where user='"&info(0)&"'"
end if

'�ڪ���O
rs.addnew
rs.movelast
rs.update "user",info(0)
rs.update "function","�ڪ���O"
rs.update "open",request.form("diary")

'�ڪ���ñ
rs.addnew
rs.movelast
rs.update "user",info(0)
rs.update "function","�ڪ���ñ"
rs.update "open",request.form("bookmark")

rs.close
conn.close
%>
<script language="Vbscript">
msgbox"�ק令�\�I",0,"����"
history.back
</script>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
