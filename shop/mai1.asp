<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=Request.Form("name")
id=Request.Form("shopname")
num=int(abs(Request.form("num")))
shang=Request.Form("shang")
money=int(abs(Request.form("money")))
money=money*num

%><!--#include file="data.asp"--><%
sql="SELECT ���D FROM �ө�"
rs1.open sql,conn1,1,1
user=rs1("���D")
if user=info(0) then 
%>
<script language="vbscript">
alert("�O�èӡA���i�H�ۤv�R��F��r�I")
history.back(1)
</script>
<%Response.End 
end if
sql="select �Ȩ� from �Τ� where id="&info(9)
set rs=conn.Execute (sql)
if rs("�Ȩ�")<money then 
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
alert("�z���Ȩ⤣���A���I���A�ӧa�I")
history.back(1)
</script>
<%Response.End 
end if
sql="select ����,�ƶq,����,���O,��O,����,���s,��T��,����,�Ȩ�,sm from �ө����~ where �֦���="&id&" and ���~�W='"&name&"'"
set rs1=conn1.execute(sql)
if rs1("�ƶq")<int(num) then
set rs1=nothing
conn1.close
set conn1=nothing
%>
<script language="vbscript">
alert("�ө��s�U�����I")
history.back(1)
</script>
<%Response.End 
end if
lx=rs1("����")
nl=rs1("���O")
tl=rs1("��O")
gj=rs1("����")
fy=rs1("���s")
dj=rs1("����")
jgd=rs1("��T��")
yin=rs1("�Ȩ�")
say=rs1("����")
sm=rs1("sm")
sql="update �Τ� set �Ȩ�=�Ȩ�-"&money&" where id="&info(9)
conn.execute(sql)
sql="update �ө� set �`�U���=�`�U���+"&money&" where �ө��W='"&shang&"'"
conn1.execute(sql)

sql="update �ө����~ set �ƶq=�ƶq-"&num&" where �֦���="&id&" and ���~�W='"&name&"'"
set rs1=conn1.execute(sql)

sql="select id from ���~ where ���~�W='"&name&"' and �֦���=" & info(9)
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
	sql="insert into ���~(���~�W,�֦���,����,����,sm,���O,��O,����,���s,��T��,�ƶq,�Ȩ�,����,�|��) values ('"&name&"',"&info(9)&",'"&lx&"','"&say&"',"&sm&","&nl&","&tl&","&gj&","&fy&","&jgd&","&num&","&yin&","&dj&","&aaao&")"
	conn.execute sql
else
id=rs("id")
	sql="update ���~ set �ƶq=�ƶq+" & num & " where id="&id
	conn.execute sql
end if
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
alert("�ʶR���\�I")
location.href = "javascript:history.go(-1)"
</script>
