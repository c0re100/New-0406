<%
set conn=server.CreateObject("adodb.connection")
path=server.MapPath("db.mdb")
conn.open "provider=microsoft.jet.oledb.4.0; data source="&path&""
%>