<%
dbname="driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("../myhome/setup/setup.asp")
set conn=server.createobject("adodb.connection")
conn.open dbname
set rs=server.createobject("adodb.recordset")
%>