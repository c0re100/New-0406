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
<!--#include file="data2.asp"-->
<!--#INCLUDE file="check.asp"-->
<%
guwst=check(request.form("guestmemo"),"�d��")
%>

<%
rs.open "select user from userinfo where user='"&info(0)&"'",conn,1,3
conn.execute "update userinfo set guestmemo='"&putmeg(request.form("guestmemo"))&"' where user='"&info(0)&"'"
rs.close
conn.close
%>
<script language="Vbscript">
msgbox"�g�J���\�I",0,"����"
history.back
</script>


<%
'=====================================================
'�ഫ���g�J�ƾڮw���榡�A�h���D�k�r��
'�եΡGputmeg(�r�Ŧ�)

function putmeg (inputmsg)
putmeg=replace(putmeg,"<",".")
putmeg=replace(inputmsg,chr(13)&chr(10),"<br>")
putmeg=replace(putmeg," ","&nbsp;&nbsp;")
putmeg=replace(putmeg,"'","��")

end function
'=====================================================

%>

<head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
