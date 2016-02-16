<%
Set conn = Server.CreateObject("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
Set rs = Server.CreateObject("ADODB.Recordset")
Set rs1 = Server.CreateObject("ADODB.Recordset")
SortSql = "Select *From news order by id Desc"
rs.Open SortSql, conn, 1,3
SortSql1 = "Select *From wed order by wed Desc"
rs1.Open SortSql1, conn, 1,3
%>