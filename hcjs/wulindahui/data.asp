<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

connstr1="DBQ="+server.mappath("fight.mdb")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn1=Server.CreateObject("ADODB.CONNECTION")
conn1.open connstr1
%>