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
<!--#INCLUDE file="check.asp"-->
<%
c1=check(request.form("name"),"�d���H")
c1=check(request.form("message"),"�d��")
%>
<%
if trim(request.form("name"))="" or trim(request.form("message"))="" then
%>
<script language="Vbscript">
msgbox "�ж�g����I",0,"����"
history.back
</script>
<%Response.End
end if

rs.open "select * from guestmessage where message='"&request.form("message")&"'",conn,1,1
if not rs.bof then
%>

<script language="Vbscript">
msgbox "�Ф��n���_��J�I",0,"����"
history.back
</script>



<%
rs.close
conn.close
Response.End
end if
rs.close
if trim(request.form("name"))=info(0) then
%>
<script language="Vbscript">
msgbox "�A���S���d���ڡA���ۤv�o����������ζܡH",0,"����"
history.back
</script>
<%
Response.End
end if

rs.open "select * from usertype where name='"&info(0)&"' and user='"&request.form("name")&"' and type='��'",conn,1,1
if not rs.bof then
%>
<script language="Vbscript">
msgbox "�z�w�g�Q�o�ӥΤ�[�J�F�¦W��A����o�e�I",0,"����"
history.back
</script>
<%
rs.close
conn.close
Response.End
end if
rs.close

rs.open "select * from userinfo where user='"&trim(request.form("name"))&"'",conn,1,1
if rs.bof then
%>

<script language="Vbscript">
msgbox "�S���o�ӥ����I",0,"����"
history.back
</script>
<%
rs.close
Response.End
end if
rs.close

if request.form("receipt")="ON" then
receipt="�O"
else
receipt="�_"
end if




rs.open "select * from guestmessage",conn,1,3
rs.addnew
rs.movelast
rs.update "send",info(0)
rs.update "receive",trim(request.form("name"))
rs.update "state","1"
rs.update "date",ccdate(date)&" "&cctime(time)
rs.update "messagetype","�@�����"
rs.update "receipt",receipt
conn.execute "update guestmessage set message='"&putmeg(request.form("message"))&"' where id="&rs("id")
%>

<script language="Vbscript">
msgbox "�z�������w�g���\�o�e���z���B�͡I",0,"����"
history.back
</script>