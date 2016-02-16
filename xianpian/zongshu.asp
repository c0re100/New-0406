<%
dim List9
dim i9
dim strSQL9
dim conn9,connstr9
connstr9="DBQ="+server.mappath("cidu-net.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
set conn9=server.createobject("ADODB.CONNECTION")
conn9.open connstr9

strSQL9="SELECT * FROM List"
set list9=server.createobject("adodb.recordset")
list9.open strSQL9,conn9,1,1
if list9.eof and list9.bof then _
Response.End
do while not (list9.eof or err)
i9=i9+1
list9.movenext
loop
list9.close
response.write i9
%>
