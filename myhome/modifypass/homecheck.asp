<%
connstr1="DBQ="+server.mappath("../setup/setup.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn1=Server.CreateObject("ADODB.CONNECTION")
conn1.open connstr1
%>
<%set rs1=conn1.execute("select house from userinfo where user='"&info(0)&"'")
if rs1.eof then
response.redirect"../wrong.htm"%>
<%rs1.close
conn1.close
else
if rs1("house")="0" then
response.redirect"wrong.htm"%>
<%
rs1.close
conn1.close
end if
end if%>