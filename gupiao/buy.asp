<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if
uname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="select �Ȩ� from �Τ� where �m�W='" & uname & "'"
set rs=conn.execute(sql)
nowmoney=rs("�Ȩ�")
conn.close
set rs=nothing %>
<!--#include file="jhb.asp"-->
<%sid=Request.Form ("sid")
ushare=abs(Request.Form ("ushare"))
if ushare<1000 then%>
<script language="vbscript">
  alert "�̧C�R�J��쬰�@��1000��"
  location.href = "index.asp"
</script>
<%else
sql= "select * from �Ѳ� where sid="&sid        
set rs=conn.execute(sql) 
if (rs("��e����")-rs("�}�L����"))/rs("�}�L����")>=0.5  then
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
  alert "���Ѳ���������R�F"
  location.href = "index.asp"
</script>


<%else
set rs=nothing

dim uname
sql="select * from �Ȥ� where �b��='"&uname&"'"
set rs=conn.execute(sql)
username=rs("�b��")
sql= "select sum(����q) As �` from ��� where �b��='"&username&"' and �ާ@='�R'"
set rs_j=conn.execute(sql) 
if  not rs_j.eof then
yi=rs_j("�`")
if yi+ushare>20000000 then 
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
  alert "���F�קK�j�᪣�@�A�ѥ��W�w����C�Ѱ����ʶR2000�U�ѪѲ��A�A���i�H�R�o��h�F"
  location.href = "index.asp"
</script>
<%

elseif ushare>200000000 then 
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
  alert "���F�קK�j�᪣�@�A�ѥ��W�w����C�Ѱ����ʶR20000�U�ѪѲ��A��X�����ƶq"
  location.href = "index.asp"
</script>
<%else
set rs_s=conn.execute("select �y�q�Ѳ�,�}�L����,��e����,���~ from �Ѳ� where sid="&sid)
set rs_u1=conn.execute("select ���Ѽ�,�b��,�R�J���� from �j�� where sid="&sid)

dsshare=rs_s("�y�q�Ѳ�")-ushare

if nowmoney>=ushare*rs_s("��e����") and dsshare>=0 then
   delmoney=ushare*rs_s("��e����")


set rs_u=conn.execute ("select ���Ѽ�,�R�J���� from �j�� where sid="&sid&" and �b��='"&username&"'")

if rs_u.eof then
sprice=rs_s("��e����")
if (sprice-rs_s("�}�L����"))/rs_s("�}�L����")>0.5 then 
sprice=rs_s("�}�L����")*1.5
elseif (sprice-rs_s("�}�L����"))/rs_s("�}�L����")<-0.5 then 
sprice=rs_s("�}�L����")*0.5
end if
sql="update �Ѳ� set ��e����="&sprice&", �y�q�Ѳ�="&dsshare&",����q=����q+"&ushare&" where sid="&sid
conn.execute sql
sql="insert into �j�� (�b��,sid,�R�J����,��������,���Ѽ�,���~,�ɶ�) values ('"&username&"','"&sid&"','"&rs_s("��e����")&"','"&rs_s("��e����")&"','"&ushare&"','"&rs_s("���~")&"','"&date()&"')"
conn.execute sql
sql="insert into ��� (�b��,sid,�ާ@,����q,�ɶ�,���~) values ('"&username&"','"&sid&"','�R','"&ushare&"','"&date()&"','"&rs_s("���~")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update �Τ� set �Ȩ�=�Ȩ�-"&delmoney&" where  �m�W='"&username&"'"
conn.execute sql
conn.close
set rs=nothing %>
<!--#include file="jhb.asp"--><%
else
yuan=rs_u("���Ѽ�")
inprice=rs_u1("�R�J����")
sprice=rs_s("��e����")
if (sprice-rs_s("�}�L����"))/rs_s("�}�L����")>0.5 then 
sprice=rs_s("�}�L����")*1.5
elseif (sprice-rs_s("�}�L����"))/rs_s("�}�L����")<-0.5 then 
sprice=rs_s("�}�L����")*0.5
end if
uprice=(rs_s("��e����")*ushare+rs_u("�R�J����")*yuan)/(ushare+yuan)
sql="update �Ѳ� set ��e����="&sprice&", �y�q�Ѳ�="&dsshare&",����q=����q+"&ushare&" where sid="&sid
conn.execute sql
sql="update �j�� set ��������="&uprice&",���Ѽ�="&rs_u("���Ѽ�")&"+"&ushare&" where �b��='"&username&"' and sid="&sid
conn.execute sql
sql="insert into ��� (�b��,sid,�ާ@,����q,�ɶ�,���~) values ('"&username&"','"&sid&"','�R','"&ushare&"','"&date()&"','"&rs_s("���~")&"')"
conn.execute sql
conn.close
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update �Τ� set �Ȩ�=�Ȩ�-"&delmoney&" where  �m�W='"&username&"'"
conn.execute sql
conn.close
end if%>
<!--#include file="jhb.asp"--><%
Randomize
sj1=int(10*rnd)+1
if sj1>4 then
sql="update �Ѳ� set ��e����=��e����*"&(1-(sj1-1)/500)&" where sid="&sid
conn.execute sql 
else
sql="update �Ѳ� set ��e����=��e����*"&(1+sj1/500)&" where sid="&sid
conn.execute sql 
end if
elseif dsshare<0 then
Response.Redirect ("serror.html")
else 
Response.Redirect ("uerror.html")
end if
end if
%>

<html>

<head>
<meta NAME="GENERATOR" Content="Microsoft FrontPage 4.0">
<link rel="stylesheet" type="text/css" href="../style1.css">
<script ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--

function button1_onclick() {
window.history.go(-2);
}

//-->
</script>

<title>�Ѳ�������\</title>
<link rel="stylesheet" href="../../CSS/STYLE.CSS">
</head>

<body bgcolor="#FFFEF4">
<table border="0" width="100%" bgcolor="#FBE651" cellspacing="0" cellpadding="0">
  <tr bgcolor="#FFFAE1"> 
    <td width="100%" height="20"> 
      <p align="center"><span class="p"><font color="#006633">�ʶR���\�I</font></span>
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











