<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Set conn1=Server.CreateObject("ADODB.CONNECTION")
Set rs1=Server.CreateObject("ADODB.RecordSet")
bbs="DBQ="+server.mappath("../bbs/bbs.asa")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};"
conn1.open bbs%>