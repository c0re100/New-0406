<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="fun.inc"-->
<%Response.Expires=0
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if
name=trim(request.form("name"))
jjr=request.form("jjr")
%>
<!--#include file="jhb.asp"-->
<%
sql="select * from �Ȥ� where �b��='"&name&"'"
set rs=conn.execute(sql)
if rs.eof then
	sql="insert into �Ȥ�(�b��,�g���H,���)values('" & name & "','"&jjr&"',0)"
	conn.execute sql
	mess="�Ѳ��b��ӽЦ��\�I�ڭ̦w�ƤF���P�g���H"&jjr&"���z�A��."
else
	sql="update �Ȥ� set �g���H='"&jjr&"' where �b��='"&name&"'"
	conn.execute sql
	mess="�ڭ̦w�ƤF�s���g���H"&jjr&"���z�A��."
end if
conn.close
%>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=jingji.asp>��^</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>