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
msgbox "這個市民已經在您的好友或黑名單中了！",0,"提示"
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
msgbox "別傻了！你自己有必要將自己加為好友或黑名單嗎？",0,"提示"
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
rs.update "send","系統消息"
rs.update "date",ccdate(date)&" "&cctime(time)
rs.update "state","1"
rs.update "messagetype","系統消息"
rs.update "receipt","否"

if request.querystring("type")="友" then
tempmessage=info(0)&"已經將您加入了好友列表！"
else
tempmessage=info(0)&"將您加入了他的黑名單中！"
end if

conn.execute "update guestmessage set message='"&tempmessage&"' where id="&rs("id")
rs.close
conn.close
%>

<script language="Vbscript">
msgbox "操作成功！",0,"提示"
history.back
</script>