<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if
if info(0)="" then
%>
<script language="vbscript">
  alert "����i�J�A�A�٨S���n��"
  location.href = "../index.asp"
</script>
<%
else
uname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="select �Ȩ� from �Τ� where �m�W='" & uname & "'"
set rs=conn.execute(sql)
nowmoney=rs("�Ȩ�")
conn.close
set rs=nothing%>
<!--#include file="jhb.asp"-->
<%sid=Request.Form ("sid")
ushare=abs(Request.Form ("ushare"))
if ushare<1000 then%>
<script language="vbscript">
  alert "�̧C��X��쬰�@��1000��"
  location.href = "index.asp"
</script>
<%else
sql= "select * from �Ѳ� where sid="&sid        
set rs=conn.execute(sql) 
if (rs("��e����")-rs("�}�L����"))/rs("�}�L����")<=-0.5  then
%>
<script language="vbscript">
  alert "���Ѳ��^�������F"
  location.href = "index.asp"
</script>

<%else
set rs=nothing
dim uname
sql="select * from �Ȥ� where �b��='"&uname&"'"
set rs=conn.execute(sql)
username=rs("�b��")
nowmoney=rs("���")

set rs_s=conn.execute ("select �}�L����,�y�q�Ѳ�,��e����,���~ from �Ѳ� where sid="&sid)
set rs_u=conn.execute ("select ���Ѽ�,�R�J����,�������� from �j�� where sid="&sid&" and  �b��='"&username&"'")

dsshare=rs_s("�y�q�Ѳ�")+ushare


addmoney=ushare*rs_s("��e����")

if rs_u.eof then 
Response.Redirect ("nserror.html")
else
dushare=rs_u("���Ѽ�")-ushare
if dushare<0 then
Response.Redirect ("userror.html")

elseif dushare=0 then
sprice=rs_s("��e����")
if (sprice-rs_s("�}�L����"))/rs_s("�}�L����")>0.5 then 
sprice=rs_s("�}�L����")*1.5
elseif (sprice-rs_s("�}�L����"))/rs_s("�}�L����")<-0.5 then 
sprice=rs_s("�}�L����")*0.5
end if
shou=(rs_s("��e����")-rs_u("��������"))*ushare
conn.execute "update �Ѳ� set ���="&date()&",����q=����q+"&ushare&",��e����="&sprice&", �y�q�Ѳ�="&dsshare&" where sid="&sid
conn.execute "delete from �j�� where �b��='"&username&"' and sid="&sid
sql="insert into ��� (�b��,sid,�ާ@,����q,�ɶ�,����,���~) values ('"&username&"','"&sid&"','��','"&ushare&"','"&date()&"','"&shou&"','"&rs_s("���~")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update �Τ� set �Ȩ�=�Ȩ�+("&addmoney&"*0.99) where  �m�W='"&username&"'"
conn.execute sql
conn.close%>
<!--#include file="jhb.asp"--><%
else
sprice=rs_s("��e����")
if (sprice-rs_s("�}�L����"))/rs_s("�}�L����")>0.5 then 
sprice=rs_s("�}�L����")*1.5
elseif (sprice-rs_s("�}�L����"))/rs_s("�}�L����")<-0.5 then 
sprice=rs_s("�}�L����")*0.5
end if
shou=(rs_s("��e����")-rs_u("��������"))*ushare
conn.execute "update �Ѳ� set ���="&date()&",����q=����q+"&ushare&",��e����="&sprice&", �y�q�Ѳ�="&dsshare&" where sid="&sid
conn.execute "update �j�� set ���Ѽ�="&dushare&" where �b��='"&username&"' and sid="&sid
sql="insert into ��� (�b��,sid,�ާ@,����q,�ɶ�,����,���~) values ('"&username&"','"&sid&"','��','"&ushare&"','"&date()&"','"&shou&"','"&rs_s("���~")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update �Τ� set �Ȩ�=�Ȩ�+("&addmoney&"*0.99) where  �m�W='"&username&"'"
conn.execute sql
conn.close
end if%>
<!--#include file="jhb.asp"--><%
Randomize
sj1=int(10*rnd)+1
if sj1>5 then
sql="update �Ѳ� set ��e����=��e����*"&(1-(sj1-1)/500)&" where sid="&sid
conn.execute sql 
else
sql="update �Ѳ� set ��e����=��e����*"&(1+sj1/500)&" where sid="&sid
conn.execute sql 
end if
end if
%>

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="../style1.css">
<title>�Ѳ�������\</title>
<link rel="stylesheet" href="../../CSS/STYLE.CSS">
</head>

<body bgcolor="#FFFEF4">
<table border="0" width="100%" cellspacing="0" cellpadding="0">
  <tr bgcolor="#FFFAE1"> 
    <td width="100%" height="20"> 
      <p align="center"><span class="p"><font color="#006633">�A�w�g���\����X�F�Ѳ��A�O�ᮬ�@�I^_^</font></span>
    </td>
  </tr>
</table>

<dl>
  <dd align="center"><p align="center"><a href="index.asp">��^�Ѳ�����j�U</a></p>
  </dd>
</dl>
</body>
</html>
<%end if
end if
end if


%>
<%conn.close
set conn=nothing
%>
