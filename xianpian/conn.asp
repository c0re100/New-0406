<%
dim conn,connstr
connstr="DBQ="+server.mappath("cidu-net.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};" 
set conn=server.createobject("ADODB.CONNECTION")
conn.open connstr
%>