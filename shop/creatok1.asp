<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
shopname=Request.Form("shopname")
memo=Request.Form("memo")
shoptype=Request.Form("shoptype")

if shopname="" or memo="" or shoptype="�п��" then 
%><script language="vbscript">
alert("�C������Ъ�^����!")
history.back(1)
</script>
<%Response.End 
end if
%><!--#include file="data.asp"--><%
rs.open "select �|������,����,�Ȩ� FROM �Τ� WHERE id="&info(9),conn
if rs("�|������")=0 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�ө��ѪO�ثe����|���}��I');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
if rs("�Ȩ�")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�A�ݭn�a��2���Ȩ�@���ө����`�U����I');location.href = 'window.close()';</script>"
	Response.End
end if
if rs("����")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�}�]�ө��ݭn�A���԰��Ŧb100�ťH�W�I');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
rs.close
rs.open "select id FROM ���~ WHERE (����='�A��' or ����='�k��' or ����='����' or ����='�ī~' or ����='�˹�' or ����='����') and  �ƶq>0 and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�}�]�ө��ݭn�A��6�إH�W���~�I');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
conn.execute "update �Τ� set �Ȩ�=�Ȩ�-200000000 where id="&info(9)
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
sql="SELECT * FROM �ө� where �ө��W='"&shopname&"' or ���D='"&info(0)&"'"
set rs1=conn1.Execute (sql)
if (not rs1.EOF) or (not rs1.BOF) then
set rs1=nothing
conn1.close
set conn1=nothing%>
<script language="vbscript">
alert("���ө��W�w�Q�`�U�ΧA�w�g�O�ѪO�F�Э��s���!")
window.close()
</script>
<%Response.End 
end if

sql="insert into �ө� (�ө��W,���D,id,����,�ө�����,�`�U���) values ('"&shopname&"','"&info(0)&"',"&info(9)&",'"&memo&"','"&shoptype&"',200000000)"
conn1.Execute(sql)
set rs1=nothing
conn1.close
set conn1=nothing
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
sd(199)="<font color=FFD7F4>�i�����j����"&info(0)&"��["&shopname&"]�ө��}�i</font>�`�U���2��....<font color=FFD7F4>�w��j�a���{�I</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
</center>
<script language="vbscript">
alert("�ӽЦ��\!")
window.close()
</script>