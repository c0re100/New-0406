<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
set rs1=server.createobject("adodb.recordset")
connstr1="DBQ="+server.mappath("setup/setup.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn1=Server.CreateObject("ADODB.CONNECTION")
conn1.open connstr1
set rs1=server.createobject("adodb.recordset")%>