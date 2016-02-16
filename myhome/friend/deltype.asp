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
rs.update "send","系統消息"
rs.update "date",ccdate(date)&" "&cctime(time)
rs.update "state","1"
rs.update "messagetype","系統消息"
rs.update "receipt","否"

tempmessage=info(0)&"已經從黑名單中撤消了您！現在您可以給他留言了！ ：）"

conn.execute "update guestmessage set message='"&tempmessage&"' where id="&rs("id")
rs.close
end if

%>


<%conn.close%>
<script language="Vbscript">
msgbox "成功執行！",0,"提示"
history.back
</script>