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

<!--#include file="../data1.asp"-->
<%
rs.open "select * from click where user='"&request.querystring("user")&"' and guest='"&info(0)&"' and postid='"&request.querystring("id")&"' and date='"&ccdate(date)&"'",conn,1,3
if not rs.bof then
rs.close
conn.close
Response.Redirect( request.querystring("url"))
%>
<%
else
rs.addnew
rs.movelast
rs.update "user",request.querystring("user")
rs.update "guest",info(0)
rs.update "postid",request.querystring("id")
rs.update "type","我的書簽"
rs.update "date",ccdate(date)
rs.close

rs.open "select * from bookmark where id="&request.querystring("id")&"",conn,1,3
number=rs("number")
if number="" then
number=1
end if
number=number+1
rs.update "number",number
rs.close

rs.open "select * from bookmark_count where user='"&request.querystring("user")&"'",conn,1,3
if rs.bof then
rs.addnew
rs.movelast
rs.update "user",request.querystring("user")
rs.update "count",1
else
rs.update "count",rs("count")+1
end if
rs.close
conn.close
Response.Redirect( request.querystring("url"))
end if
%>

<%
'=====================================================
'轉換日期型或字符日期型為全格式如：2000-01-01
'調用：ccdate(日期變量或日期格式字符串)

function ccdate(sdate)
temp=cdate(sdate)
if len(month(temp))=1 then
m="0"&month(temp)
else
m=month(temp)
end if

if len(day(temp))=1 then
d="0"&day(temp)
else
d=day(temp)
end if
ccdate=year(temp)&"-"&m&"-"&d
end function
'=====================================================
%>