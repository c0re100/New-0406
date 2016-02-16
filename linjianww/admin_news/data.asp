<%
Set Conn = Server.CreateObject("ADODB.Connection")
jiutiandata="DBQ="+server.mappath("../../news.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open jiutiandata
%>