<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0

if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
sheepname=request.form("sheepname")
if sheepname="" then
response.redirect"Warning.htm"
else
%>
<!--#include file="data1.asp"-->
<%
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'",conn,1,1
if rs.bof then
rs.close
cn.close
response.redirect"Warning.htm"
else
rs.close
rs.open"select * from rules ",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"Warning.htm"
else
eatdel=rs("eatdel")
eatplushungry=rs("eatplushungry")
eatplusmilk=rs("eatplusmilk")
rs.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ� from �Τ� where id="&info(9),conn,1,1
tempsplosh=rs("�Ȩ�")-eatdel
if tempsplosh<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�z�S�����������I���h���I���a�C",0,"FLUSH"
history.back
</script>
<%
else
rs.close%>
<!--#include file="data1.asp"-->
<%
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'",conn,1,1
feeddate=rs("feeddate")
workload=rs("workload")
hungry=rs("hungry")+eatplushungry
if hungry>100 then
hungry=100
end if
milk=rs("milk")+eatplusmilk
rs.close
if datediff("d",feeddate,date)<>0 then 
tempdate=date
conn.execute"update sheep set milk='"&milk&"',hungry='"&hungry&"',feeddate='"&date&"',feedsheepday=feedsheepday+1,workload='1' where sheepname='"&sheepname&"' and user='"&info(0)&"'"

conn.close
%>
<script language="vbscript">
msgbox"�޾i�����I",0,"FLUSH"
history.back
</script>
<%
rs.close
conn.close
else
rs.open"select * from sheep where sheepname='"&sheepname&"' and user='"&info(0)&"'",conn,1,1

if workload>=3 then
conn.close
%>
<script language="vbscript">
msgbox"�z�w�g���@�p�ĤT���I���ѦA�ӧa�C:-)",0,"FLUSH"
history.back
</script>
<%
else
conn.execute"update sheep set milk='"&milk&"',hungry='"&hungry&"',workload=workload+1 where sheepname='"&sheepname&"' and user='"&info(0)&"'"
rs.close
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute"update �Τ� set �Ȩ�='"&tempsplosh&"' where id="&info(9)
conn.close
%>
<script language="vbscript">
msgbox"�޾i�����I",0,"FLUSH"
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