
<body bgcolor="#660000">
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(3)<30 then Response.Redirect "../error.asp?id=460"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from �Τ� where �m�W='"&info(0)&"' and ����='�C�L'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�C�L�����H���i�H�i�榹�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select �x�� from ���� where �x��='"&info(0)&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�w�g�O�x���F,�٬D�Խ֩O�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select �x�� from ���� where ����='"&info(5)&"'",conn
zhangmen=rs("�x��")
rs.close
rs.open "select ����,�Z�\,�y�O,���� from �Τ� where �m�W='"&zhangmen&"'",conn
dj=rs("����")
wg=rs("�Z�\")
ml=rs("�y�O")
zz=rs("����")
rs.close
rs.open "select ����,����,�Z�\,�y�O,���� from �Τ� where id="&info(9),conn
if rs("����")<dj or rs("�Z�\")<wg or rs("�y�O")<ml or rs("����")<zz then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�����šB�Z�\�B�y�O�B���賣��W�V�A�����x���ܡH�^�h�A�n�n�m�m�a�I');window.close()}</script>"
	Response.End 
end if
conn.execute("update �Τ� set ����='�̤l',grade=1 where �m�W='"&zhangmen&"'")
conn.execute("update �Τ� set ����='�x��',grade=5 where id="&info(9))
conn.execute("update ���� set �x��='"&info(0)&"' where ����='"&info(5)&"'")
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�D�Դx���j"& info(0) &"</font><font color=FFD7F4>�H�u�q���Z�\�W�V�F�����x������ "&info(5)&" ���s�x���I</font></font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{alert('�A�H�u�q���Z�\���������s���x��,�аh�X���򭫷s�i�h�I');window.close();}</script>"
Response.End
%>