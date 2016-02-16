<%@ LANGUAGE = VBScript %>
<!--#include file="data.asp" -->
<!--#include file="../inc/global.news.asp" -->
<!--#include file="../inc/function.inc.asp" -->
<%
action=Request("action")
id=trim(Request.Form("id"))
news_type=trim(Request.Form("news_type"))
title=trim(Request.Form("title"))
content=trim(Request.Form("content"))
author=trim(Request.Form("author"))
hot_sign=trim(Request.Form("hot_sign"))
if (hot_sign="") then
hot_sign="0"
end if
news_form=trim(Request.Form("news_form"))
filltime_year=Request.Form("filltime_year")
filltime_month=Request.Form("filltime_month")
filltime_day=Request.Form("filltime_day")
filltime=filltime_year&"-"&filltime_month&"-"&filltime_day&" 01:01:01"
<!--#include file="data.asp" -->
if action="add" then
ErrMsg=""
CheckValue "新聞標題",title,0,1,250
CheckValue "新聞內容",content,0,1,99999999
if (ErrMsg="") then
sql_field=array("title","content","type","author","hot_sign","news_form","filltime")
sql_values=array(title,content,news_type,author,hot_sign,news_form,filltime)

'		sql_field=array("author","news_form")'
'		sql_values=array(author,news_form)

for i=0 to ubound(sql_field)
sql_values(i)=replace(sql_values(i),"'","''")
next
sql = "INSERT INTO hc_news (" & Join(sql_field, ",") & ") VALUES ('" & Join(sql_values, "','") & "')"
'Response.Write sql
'Response.end
conn.Execute sql
word="你的信息已成功送出!"
call MsgBox(word)
else
call Msgout(ErrMsg,"javascript:window.history.back()",3)
end if
conn.close
set conn=nothing
end if
if action="modify" then
id=Request.Form("id")

ErrMsg=""
CheckValue "新聞標題",title,0,1,250
CheckValue "新聞內容",content,0,1,99999999

if (ErrMsg="") then
sql_field=array("title","content","type","author","hot_sign","news_form","filltime")
sql_values=array(title,content,news_type,author,hot_sign,news_form,filltime)

sql="UPDATE hc_news SET "
for i=0 to ubound(sql_field)
sql = sql & sql_field(i) & "='" & replace(sql_values(i),"'","''") & "'"
if i <> ubound(sql_field) then
sql=sql & ","
else
'如果你的主關鍵字名不是"id",請修改下面的代碼
sql=sql & " where id="&id
end if
next
conn.Execute sql
word="你的信息已成功送出!"
call MsgBox(word)
else
call Msgout(ErrMsg,"javascript:window.history.back()",3)
end if

conn.close
set conn=nothing

end if


if (action="del_row") then
id=Request.Querystring("id")

sql="DELETE FROM hc_news where id="&id
conn.Execute sql
conn.close
set conn=nothing
word="你的信息已成功刪除!"
call MsgBox(word)
end if
%>