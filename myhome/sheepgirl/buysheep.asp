<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")


function ccdate(sdate)
temp=cdate(sdate)
if len(month(temp))=1 then
m="0"&month(temp)
else
m=month(temp)
end if

if len(day(temp))=1 then
d="0"&day(temp)
else
d=day(temp)
end if
ccdate=year(temp)&"-"&m&"-"&d
end function

function cctime (stime)
if len(hour(stime))=1 then
h="0"&hour(stime)
else
h=hour(stime)
end if
if len(minute(stime))=1 then
m="0"&minute(stime)
else
m=minute(stime)
end if
if len(second(stime))=1 then
s="0"&second(stime)
else
s=second(stime)
end if
cctime=h&":"&m&":"&s
end function
'=====================================================
sheepname=request.form("sheepname")
if sheepname="" or len(sheepname)>16 then
response.redirect"warning.htm"
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open"select �Ȩ� from �Τ� where �m�W='"&info(0)&"'",conn,1,1
if rs("�Ȩ�")<100000 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�z�S�����������I",0,"FLUSH"
history.back
</script>
<%
else
money=rs("�Ȩ�")-100000
conn.execute("update �Τ� set �Ȩ�='"&money&"'  where �m�W='"&info(0)&"'")
rs.close%>
<!--#include file="data1.asp"-->

<%rs.open"select user from sheep where user='"&info(0)&"'",conn,1,1
if not rs.bof then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�@�ӤH�����i�@�Ӥp��",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select user from sheep",conn,1,1
if rs.recordcount>10000 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"�����̭����Ϥw�g�W�L�F10000���A�A����A��i",0,"FLUSH"
history.back
</script>
<%
else
rs.close
rs.open"select cleaninit,happyinit,healthinit,milkinit,lifeinit,hungryinit from rules",conn,1,1
cleaninit=rs("cleaninit")
sheephappyinit=rs("happyinit")
sheephealthinit=rs("healthinit")
milkinit=rs("milkinit")
lifeinit=rs("lifeinit")
hungryinit=rs("hungryinit")
sheepdate=date
feeddate=DATEADD("d",-1,DATE)
workload=0
rs.close
conn.execute"insert into sheep(sheephappy,sheephealth,user,sheepname,sheepdate,life,milk,hungry,workload,clean,feeddate,feedsheepday,logintoday) values('"&sheephappyinit&"','"&sheephealthinit&"','"&info(0)&"','"&sheepname&"','"&sheepdate&"','"&lifeinit&"','"&milkinit&"','"&hungryinit&"','"&workload&"','"&cleaninit&"','"&feeddate&"','0','"&feeddate&"')"
rs.open"select id from sheep",conn,1,1
rs.movelast
sheepid=rs("id") '�p�Ͻs��
conn.execute"update sheep set user='"&info(0)&"' where sheepname='"&sheepname&"'"

%>


<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
</div>
<div align="center">
  <center>
  <table border="0" width="470" bordercolor="#FFFFFF" cellspacing="1"
  cellpadding="0" height="20">
    <tr>
        <td class="p6"><font size="2">��-�z��e����m---<a
        href="indexsheep.asp">�t��|</a>-<a
        href="feedsheep.asp">�����</a></font></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
  <table border="0" width="470" cellspacing="1" cellpadding="0" height="50">
    <tr>
      <td class="p2" width="100%">
        <p align="center"><font size="2">���ߡI�z�w�g���\��i�F<%=sheepname%>�A�z��i���Ĥl���s���O�G</font><%=rs("id")%></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <table border="0" width="470" cellspacing="1" cellpadding="0">
    <tr>
      <td class="p3" width="100%">
        <p align="right"><font size="2">�֦^����Ҷ}�l���U<%=sheepname%>�a�I</a></font></td>
    </tr>
  </table>
</div>
<div align="center">
  <center>
    <font color="#FFFFFF"><a href="javascript:window.close()">�������f</a></font>
</center>
</div>

</body>

</html>
<%
rs.close 
conn.close 
end if 
end if
end if 
end if %>
