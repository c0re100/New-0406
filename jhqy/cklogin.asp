<!--#include file="config.asp"-->
<%
varadm=request.form("username")
varpwd=request.form("password")

if varadm="" then
      response.redirect "error.asp?msg=�п�J�޲z���W�r!"
end if
if varpwd="" then
      response.redirect "error.asp?msg=�п�J�޲z���K�X!"
end if

if varadm=adm and varpwd=pwd then
      session("admin")="ture"
      response.redirect "manage.asp"
else
      response.redirect "error.asp?msg=�A���Τ�W�M�K�X�����D!"
end if
%>