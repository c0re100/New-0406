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
c1=check(request.form("name"),"留言人")
c1=check(request.form("message"),"留言")
%>
<%
if trim(request.form("name"))="" or trim(request.form("message"))="" then
%>
<script language="Vbscript">
msgbox "請填寫完整！",0,"提示"
history.back
</script>
<%Response.End
end if

rs.open "select * from guestmessage where message='"&request.form("message")&"'",conn,1,1
if not rs.bof then
%>

<script language="Vbscript">
msgbox "請不要重復輸入！",0,"提示"
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
msgbox "你有沒有搞錯啊，給自己發消息有什麼用嗎？",0,"提示"
history.back
</script>
<%
Response.End
end if

rs.open "select * from usertype where name='"&info(0)&"' and user='"&request.form("name")&"' and type='黑'",conn,1,1
if not rs.bof then
%>
<script language="Vbscript">
msgbox "您已經被這個用戶加入了黑名單，不能發送！",0,"提示"
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
msgbox "沒有這個市民！",0,"提示"
history.back
</script>
<%
rs.close
Response.End
end if
rs.close

if request.form("receipt")="ON" then
receipt="是"
else
receipt="否"
end if




rs.open "select * from guestmessage",conn,1,3
rs.addnew
rs.movelast
rs.update "send",info(0)
rs.update "receive",trim(request.form("name"))
rs.update "state","1"
rs.update "date",ccdate(date)&" "&cctime(time)
rs.update "messagetype","一般消息"
rs.update "receipt",receipt
conn.execute "update guestmessage set message='"&putmeg(request.form("message"))&"' where id="&rs("id")
%>

<script language="Vbscript">
msgbox "您的消息已經成功發送給您的朋友！",0,"提示"
history.back
</script>