<%
Set Conn=Server.CreateObject("ADODB.Connection")
set rs=server.createobject("adodb.recordset")
Connstr="DBQ="+server.mappath("../setup/setup.asp")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
conn.open Connstr
%>
