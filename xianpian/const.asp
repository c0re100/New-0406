<%
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="DBQ="+server.mappath("srsr.asp")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};DriverId=36;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=6;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
Conn.Open connstr
%>

