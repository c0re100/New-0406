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


<%
rs.open "select * from usertype where user='"&info(0)&"' and name='"&request.querystring("name")&"'",conn,1,1
if not rs.bof then
%>

<script language="Vbscript">
msgbox "�o�ӥ����w�g�b�z���n�ͩζ¦W�椤�F�I",0,"����"
history.back
</script>
<%
rs.close
conn.close
Response.End
end if
if info(0)=request.querystring("name") then
%>
<script language="Vbscript">
msgbox "�O�̤F�I�A�ۤv�����n�N�ۤv�[���n�ͩζ¦W��ܡH",0,"����"
history.back
</script>
<%rs.close
conn.close
Response.End
end if%>
<%
rs.close
rs.open "select * from usertype",conn,1,3
rs.addnew
rs.movelast
rs.update "user",info(0)
rs.update "name",request.querystring("name")
rs.update "type",request.querystring("type")
rs.close

rs.open "select * from guestmessage",conn,1,3
rs.addnew
rs.movelast
rs.update "receive",request.querystring("name")
rs.update "send","�t�ή���"
rs.update "date",ccdate(date)&" "&cctime(time)
rs.update "state","1"
rs.update "messagetype","�t�ή���"
rs.update "receipt","�_"

if request.querystring("type")="��" then
tempmessage=info(0)&"�w�g�N�z�[�J�F�n�ͦC��I"
else
tempmessage=info(0)&"�N�z�[�J�F�L���¦W�椤�I"
end if

conn.execute "update guestmessage set message='"&tempmessage&"' where id="&rs("id")
rs.close
conn.close
%>

<script language="Vbscript">
msgbox "�ާ@���\�I",0,"����"
history.back
</script>