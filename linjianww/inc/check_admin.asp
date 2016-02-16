<%
if session("jiutian_grade")< 9 or Session("ADMIN")="" then
response.redirect "../login.asp"
end if

%>