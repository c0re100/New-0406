<%
if Session("jiutian_name")="" or Session("jiutian_id")="" then 
       url("../login.asp")
	response.end
end if

If Session("ADMIN") <> True Then 
  	url("../login.asp")
	response.end
end if

if Session("jiutian_grade")<9 then
url("../login.asp")
	response.end
end if 
%>