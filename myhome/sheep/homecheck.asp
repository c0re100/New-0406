<%
connstr1="DBQ="+server.mappath("../setup/setup.asp")+";DefaultDir=;DRIVER={Microsoft Access Driver (*.mdb)};"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open connstr1
%>
<%set rs=conn.execute("select house from userinfo where user='"&info(0)&"'")
if rs.eof then
response.redirect"../wrong.htm"%>
<%rs.close
conn.close
else
if rs("house")="0" then
response.redirect"wrong.htm"%>
<%
rs.close
conn.close
end if
end if%>