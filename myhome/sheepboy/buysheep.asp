<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
nickname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �t��,�ʧO,�Ȩ� from �Τ� where id="&info(9),conn
peier=rs("�t��")
if peier="" or peier="�L" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�٨S���t���ڡI');location.href = 'indexsheep.asp';}</script>"
Response.End
end if
if rs("�ʧO")="�k" then 
  sheep00a1=nickname
  sheep00a2=peiper
  else
   sheep00a1=peiper
   sheep00a2=nickname
end if
if rs("�Ȩ�")>10000000 then 
  sheepxb="�k"
 else
  sheepxb="�k"
end if


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
sheepname2=request.form("sheepname2")
if sheepname="" or len(sheepname)>16 or sheepname2="" or len(sheepname)>12 then
response.redirect"warning.htm"
else
rs.close
rs.open"select �Ȩ� from �Τ� where �m�W='"&info(0)&"'",conn,1,1
if rs("�Ȩ�")<10 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�S���ڡI');location.href = 'indexsheep.asp';}</script>"
Response.End
else
money=rs("�Ȩ�")-10
'conn.execute("update �Τ� set �Ȩ�='"&money&"'  where �m�W='"&info(0)&"'")
rs.close%>
<!--#include file="data1.asp"-->

<%rs.open"select user from sheep where user='"&info(0)&"'",conn,1,1
if not rs.bof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�p���ͨ|�Ȯɰ���ͤ@�ӫĤl�I');location.href = 'indexsheep.asp';}</script>"
Response.End
else
rs.close
rs.open"select user from sheep",conn,1,1
if rs.recordcount>10000 then
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�Ӧ���|��|�̤w�g���W�L�F10000�p�ĤF�A�A�̭n�Q�ͤp�Ī��ܡA�򯸪��pô�a�A����');location.href = 'indexsheep.asp';}</script>"
Response.End
%>

<!--#include file="data1.asp"-->
<%
else
rs.close
rs.open"select cleaninit,happyinit,healthinit,milkinit,lifeinit,hungryinit,wugonginit,neiliinit,gongjiinit from rules",conn,1,1
cleaninit=rs("cleaninit")
sheephappyinit=rs("happyinit")
sheephealthinit=rs("healthinit")
milkinit=rs("milkinit")
lifeinit=rs("lifeinit")
hungryinit=rs("hungryinit")
wugonginit=rs("wugonginit")
neiliinit=rs("neiliinit")
gongjiinit=rs("gongjiinit")
sheepdate=date
feeddate=DATEADD("d",-1,DATE)
workload=0
rs.close
conn.execute"insert into sheep(sheephappy,sheephealth,sheep001,sheepname,sheepdate,life,milk,wugong,leinl,gongji,hungry,workload,clean,feeddate,feedsheepday,logintoday) values('"&sheephappyinit&"','"&sheephealthinit&"','"&sheep00a1&"','"&sheepname&"','"&sheepdate&"','"&lifeinit&"','"&milkinit&"','"&wugonginit&"','"&neiliinit&"','"&gongjiinit&"','"&hungryinit&"','"&workload&"','"&cleaninit&"','"&feeddate&"','0','"&feeddate&"')"
rs.open"select id from sheep",conn,1,1
rs.movelast
sheepid=rs("id") '�p�Ͻs��
conn.execute"update sheep set user='"&info(0)&"' where sheepname='"&sheepname&"'"
conn.execute"update sheep set sheep002='"&sheepname2&"' where sheepname='"&sheepname&"'"
conn.execute"update sheep set sheepxb='"&sheepxb&"' where sheepname='"&sheepname&"'"
 Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

rs.open "select id from �Τ� where id="&info(9),conn
conn.execute("update �Τ� set �Ȩ�=�Ȩ�-1000000,�p��='��',�ĦW='"&sheepname&"'  where �m�W='"&info(0)&"'")
conn.execute("update �Τ� set �Ȩ�=�Ȩ�-1000000,�p��='��',�ĦW='"&sheepname&"'  where �m�W='"&peier&"'")
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0" background="../../images/8.jpg">
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
          <%if sheepxb="�k" then
		  msg="<font color=FF0000>�i�ͤp�ġj</font>����:"&info(0)&"�M"&peier&"���\���ͤU�@��<img src='../Myhome/SHEEPBOY/baby3.gif'>�Ĥl���W�r�O"&sheepname&"...."

		 		   %>
		            <p align="center"><font size="2">���ߡI�A�̦��\���ͤU�@��<img src="baby3.gif" width="144" height="133">�A�A�̫Ĥl������s���O�G</font><%=rs("id")%> 
            <%
  else
   msg="<font color=FF0000>�i�ͤp�ġj</font>����:"&info(0)&"�M"&peier&"���\���ͤU�@��<img src='../Myhome/SHEEPBOY/girl15.gif'>�Ĥl���W�r�O"&sheepname&"...."

  %>
          <p align="center"><font size="2">���ߡI�A�̦��\���ͤU�@��<img src="girl15.gif" width="45" height="57">�A�A�̫Ĥl������s���O�G</font><%=rs("id")%> 
            <%end if%>
          <p align="center">&nbsp;
        </td>
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
	set rs=nothing	
	conn.close
	set conn=nothing
end if
end if 
end if 
end if 
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>