<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="jhb.asp"--><%
jin=request.form("money")
jin=abs(int(jin))
mess=""
Jname=info(0)
sql="select * from �Ȥ� where �b��='" & Jname & "' "
set rs=conn.execute(sql)
yin=rs("���")
if jin>yin then
mess="�A���Ѳ��g���H�j�n���A�ڨS�����A�ȳo��h���r�H"
else
sql="update �Ȥ� set ���=���-"& jin & " where �b��='" & Jname & "'"
conn.execute sql
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ� from �Τ� where id="&info(9),conn
money=rs("�Ȩ�")
moneyok=int(money)+abs(fn1)
if moneyok>2000000000 then
	Response.Write "<script language=JavaScript>{alert('�A���Ȩ�N�ֶW�L�F20���F�A���������{���бz�֨��ǧa�I');location.href = 'javascript:history.go(-1)';}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sql="update �Τ� set �Ȩ�=�Ȩ�+1 where �m�W='" & Jname & "'"
conn.execute sql
mess="�A�q�Ѳ��b����X�F1��A���n�F��"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>

<head>
<title></title>
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=qu.asp>��^</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>