<%
set connt=server.createobject("adodb.connection")
path="dbq="+server.mappath("yhy.mdb")+";driver={microsoft access driver (*.mdb)};"
connt.open path
%>