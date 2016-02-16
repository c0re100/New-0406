<%
info=Session("info")
if info(0)="" then
response.redirect"warning.htm"
Response.End
end if


dbname="driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("setup/setup.asp")
set conn=server.createobject("adodb.connection")
conn.open dbname,"admin","chaopiaormb"
set rs=server.createobject("adodb.recordset")
%>