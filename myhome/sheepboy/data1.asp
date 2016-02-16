<%


dbname="driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("setup.mdb")
set conn=server.createobject("adodb.connection")
conn.open dbname,"admin","chaopiaormb"
set rs=server.createobject("adodb.recordset")
%>