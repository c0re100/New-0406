<!--#include file="config.asp"-->
<%
varadm=request.form("username")
varpwd=request.form("password")

if varadm="" then
      response.redirect "error.asp?msg=請輸入管理員名字!"
end if
if varpwd="" then
      response.redirect "error.asp?msg=請輸入管理員密碼!"
end if

if varadm=adm and varpwd=pwd then
      session("admin")="ture"
      response.redirect "manage.asp"
else
      response.redirect "error.asp?msg=你的用戶名和密碼有問題!"
end if
%>