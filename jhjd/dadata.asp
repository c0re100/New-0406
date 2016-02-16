<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("jiudian.asa")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>