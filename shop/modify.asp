<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="data.asp"--><%
name=Request.Form("name")
id=int(Request.Form("aaa"))
sql="select ����,����,���O,��O,����,���s,��T��,�Ȩ�,�ƶq,���� from ���~ where �֦���="&id&" and ���~�W='"&name&"' and �˳�=false and �ƶq>0"
set rs=conn.execute(sql)
if  rs.EOF or rs.bof then
set rs=nothing
conn.close
set conn=nothing
%>
<script language="vbscript">
alert("�S���Ӫ��~!")
history.back(1)
</script>
<%Response.End 
end if
lx=rs("����")
say=rs("����")
nl=rs("���O")
tl=rs("��O")
gj=rs("����")
fy=rs("���s")
dj=rs("����")
jgd=rs("��T��")
yin=rs("�Ȩ�")
sl=rs("�ƶq")
sql="delete * from  ���~ where �֦���="&id&" and ���~�W='"&name&"'"
set rs=conn.execute(sql)
sql="select id from �ө����~ where �֦���="&id&" and ���~�W='"&name&"' and �ƶq>0"
rs1.Open sql,conn1,1,1
if  rs1.EOF or rs1.bof then
sql="insert into �ө����~(���~�W,�֦���,����,����,���O,��O,����,���s,��T��,�ƶq,�Ȩ�,����) values ('"&name&"',"&info(9)&",'"&lx&"','"&say&"',"&nl&","&tl&","&gj&","&fy&","&jgd&","&sl&","&yin&","&dj&")"
conn1.execute sql
else
id=rs1("id")
	sql="update �ө����~ set �ƶq=�ƶq+" & sl & " where id="&id
	conn1.execute sql
end if

set rs1=nothing
conn1.close
set conn1=nothing
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "man.asp"
%>

