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

<!--#include file="../data2.asp"-->

<%
rs.open "select * from diarymessage where id="&request.querystring("id"),conn,1,1
user=rs("user")
rs.close



rs.open "select * from click where user='"&user&"' and guest='"&info(0)&"' and postid='"&request.querystring("id")&"' and type='日記'",conn,1,1


if (not rs.bof) or user=info(0) then
dbclick=1
else
dbclick=0
end if
rs.close


if dbclick=0 then
rs.open "select * from diarymessage where id="&request.querystring("id"),conn,1,3
tempclick=rs("clicknumber")+1
rs.update "clicknumber",tempclick
rs.close

rs.open "select * from click",conn,1,3
rs.addnew
rs.movelast
rs.update "user",user
rs.update "guest",info(0)
rs.update "postid",request.querystring("id")
rs.update "type","日記"
rs.update "date",ccdate(date)
rs.close
end if

Response.redirect "read_diary.asp?id="&request.querystring("id")

%>