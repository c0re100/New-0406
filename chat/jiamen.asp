<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
men1=trim(Application("menpai"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �������,�A�X FROM ���� WHERE ����='" & men1 & "'",conn
if rs("�������")<15000000  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�]����������Ӥ�,�Ȯɰ���ۦ��s�̤l�A�ФU���a�I');}</script>"
	Response.End
end if
shihe=trim(rs("�A�X"))
rs.close
rs.open "select ����,���� FROM �Τ� WHERE id="&info(9),conn
if rs("����")="�x��" or rs("����")=men1 or rs("����")="�x��" then
	Response.Write "<script language=JavaScript>{alert('�ݥJ�ӤF���n���I�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if shihe<>info(1) and shihe<>"�Ҧ�" then
	Response.Write "<script language=JavaScript>{alert('�Ӫ������ۦ�"&shihe&"�̤l�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

conn.execute "update ���� set �̤l��=�̤l��-1 where ����='" & info(5) & "'"
conn.execute "update ���� set �̤l��=�̤l��+1 where ����='" & men1 & "'"
if rs("����")<>"�C�L" then
conn.execute "update �Τ� set �O�@=false,�Z�\=100,���O=100,�Ȩ�=0,�s��=0 where id="&info(9)
jiamen="["&info(0)&"]�]��["&men1&"]���x�����e���G�ڨӰ�,�ֵ��ڿ���,�ڭn����..."&info(0)&"�I�q�F��Ӫ�["&info(5)&"]����,���\���[�J["&men1&"]!"
else
conn.execute "update �Τ� set �Ȩ�=0,�s��=0 where id="&info(9)
jiamen="["&info(0)&"]�]��["&men1&"]���x�����e���G�ڭn�[�J�A�̪����A�ֵ����I���ڡA�ڧֽa���F,"&info(0)&"���\���[�J["&men1&"]!"
end if
conn.execute "update ���� set �������=�������-10000000 where ����='" & men1 & "'"
conn.execute "update �Τ� set ����='" & men1 & "', ����='�̤l',grade=1 where id="&info(9)
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
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
sd(195)=info(0)
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i���ۤJ���j</font>"&jiamen
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
