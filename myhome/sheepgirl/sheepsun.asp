<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")

sheepname=request.form("sheepname")
if sheepname="" then
response.redirect"Warning.htm"
else
%>

<!--#include file="data2.asp"-->
<%
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
rs.close
rs.open"select * from rules ",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
sundel=rs("sundel")
sunplushealth=rs("sunplushealth")
sunplusmilk=rs("sunplusmilk")
rs.close
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open"select �Ȩ� from �Τ� where id="&info(9),conn,1,1
tempsplosh=rs("�Ȩ�")-sundel
if tempsplosh<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�z�S�����������I���h���I���a�C",0,"FLUSH"
history.back
</script>
<%
else%>
<!--#include file="data1.asp"-->
<%
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'",conn,1,1
feeddate=rs("feeddate")
workload=rs("workload")
sheephealth=rs("sheephealth")+sunplushealth
if sheephealth>100 then
sheephealth=100
end if
milk=rs("milk")+sunplusmilk
rs.close
if datediff("d",feeddate,date)<>0 then 
tempdate=date
conn.execute"update sheep set milk='"&milk&"',sheephealth='"&sheephealth&"',feeddate='"&date&"',feedsheepday=feedsheepday+1,workload='1' where sheepname='"&sheepname&"' and user='"&info(0)&"'"
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute"update �Τ� set �Ȩ�='"&tempsplosh&"' where id="&info(9)
%>
<script language="vbscript">
msgbox"����D�����I",0,"FLUSH"
history.back
</script>
<%
else
if workload>=3 then
conn.close
%>
<script language="vbscript">
msgbox"�z�w�g���@�p�ĤT���I���ѦA�ӧa�C:-)",0,"FLUSH"
history.back
</script>
<%
else
conn.execute"update sheep set milk='"&milk&"',sheephealth='"&sheephealth&"',workload=workload+1 where sheepname='"&sheepname&"' and user='"&info(0)&"'"
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute"update �Τ� set �Ȩ�='"&tempsplosh&"' where id="&info(9)
conn.close
%>
<script language="vbscript">
msgbox"����D�����I",0,"FLUSH"
history.back
</script>
<%
end if
end if
end if
end if
end if
end if
%>