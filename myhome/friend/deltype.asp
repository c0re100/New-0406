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
rs.open "select * from usertype where id="&request.querystring("id"),conn,1,3
blackname=rs("name")

rs.delete
rs.close


if request.querystring("type")<>"" then

rs.open "select * from guestmessage",conn,1,3
rs.addnew
rs.movelast
rs.update "receive",blackname
rs.update "send","�t�ή���"
rs.update "date",ccdate(date)&" "&cctime(time)
rs.update "state","1"
rs.update "messagetype","�t�ή���"
rs.update "receipt","�_"

tempmessage=info(0)&"�w�g�q�¦W�椤�M���F�z�I�{�b�z�i�H���L�d���F�I �G�^"

conn.execute "update guestmessage set message='"&tempmessage&"' where id="&rs("id")
rs.close
end if

%>


<%conn.close%>
<script language="Vbscript">
msgbox "���\����I",0,"����"
history.back
</script>