<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
Response.Buffer=true
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(2)=0 then
Response.Write "<script Language=Javascript>alert('�Ǫ��ЬO����}�ҫO�@���I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM �q�r�O WHERE �q�r�H��='" & info(0) & "' and �����f��='���'",conn
if not(rs.bof or rs.eof)  then	
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�O�b�k�q��ǡA����}�O�@�I');}</script>"
	Response.End
end if
rs.close
rs.open "select top 1 �m�W1,�m�W2 FROM ��D ",conn
if rs("�m�W1")=info(0) or rs("�m�W2")=info(0)  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�{�b�b��D���A�A���o�k�]�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
rs.open "SELECT ���H��,�_��,�O�@,�ާ@�ɶ�,�|������ FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
aaay=rs("�|������")
if sj<60 and aaay=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if

if rs("���H��")>=30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�@�ɤѤU�a��,�ٷQ�O�@�H�I');}</script>"
	Response.End
end if
if rs("�_��")="�F�C������" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�{�b�������_��[�F�C������]����i��m�\�O�@�I');}</script>"
	Response.End
end if
if rs("�O�@")=true then
	conn.Execute "update �Τ� set �O�@=false,�ާ@�ɶ�=now() where id="&info(9)	
	diaox="�\���X���A300�~�e���Z�L�G�פS�N�A�����t�F�I"
else
	conn.Execute "update �Τ� set �O�@=true,�ާ@�ɶ�=now() where id="&info(9)
	diaox="���F�׷ҵL�W���\�A�A�������A�Z�L���`��i�H�ӥ��@�}�F�I"
end if
conn.close
set rs=nothing
set conn=nothing
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
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="��"
sd(199)="<font color=#9FDF70>�i�����j" & info(0) & ""& diaox &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd

%>
<script language=javascript>
parent.f4.location.href='f4.asp';
</script> 