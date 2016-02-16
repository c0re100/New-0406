<%
id = Request("id")
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
SET rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From admin WHERE (id= " & id & ")",conn,1,3
CONN.EXECUTE "DELETE From admin WHERE id= " & id & ""
SET CONN = NOTHING
Response.Redirect ("use.asp")
%>
