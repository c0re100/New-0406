<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<8  then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim conn,rs,rsuser
on error resume next
opt=Request.QueryString("opt")
st=Request.QueryString("touser")
set rs=server.createobject("adodb.recordset")
rs.open ("select * from �Τ� where �m�W='"&st&"'"),conn,0,1
if rs.bof or rs.eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Redirect "../error.asp?id=446"
else
grade1=rs("grade")
stt=rs("id")
if info(2)<=grade1 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
alert "�藍�_�A���O�O���ź޲z�H���A�A���޲z���Ť���C�I"
window.close()
</script>
<%
response.end
end if
if st="�����`��" and st="�b�b"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
alert "�藍�_�A���O�S�j�H���I"
window.close()
</script>
<%
response.end
end if
 if opt="�e��" then
call boot(st)
conn.execute("update �Τ� set ���A='��',�n��=now() where id="&stt)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & st & "','�e��',now(),'�S�O�޲z���')"
msg="<font color=FF0000>�i�S�O�޲z��ʡj</font>�ѪŤ���M���X�@������A�⤣ť�ܪ��p��<font color=FFC01F>"&st&"</font>��J�F�x���j�c�C(�e��)[����H��="&info(0)&"]"
end if
if opt="�J��" then
conn.execute("update �Τ� set ���A='�c',�n��=now()+1/84 where id="&stt)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & st & "','���c',now(),'�S�O�޲z���')"
msg="<font color=FF0000>�i�S�O�޲z��ʡj</font>�H�ۤ@�n�ªZ���[�ܡA�x�����H�p�T���ꪺ�N��ť�ܪ��p��<font color=FFC01F>"&st&"</font>���i�F���Ǥj�c�C(���c)[����H��="&info(0)&"]"

call boot(st)
end if
if opt="��H" then
'conn.execute("update �Τ� set ���A='�c',�n��=now()+1/84 where id="&stt)
'conn.execute("update �Τ� set ���A='���`',�n��=now() where id="&stt)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & st & "','��H',now(),'�S�O�޲z���')"
msg="<font color=FF0000>�i��H�j</font>��ť�y���@�n�A<font color=FFC01F>"&st&"</font>�Q��X��ѫǡC[����H��="&info(0)&"]"
call boot(st)
end if
if opt="�����{�D" then
conn.execute("update �Τ� set ���A='�c',�n��=now()+1/4 where id="&stt)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & st & "','���c',now(),'�S�O�޲z���')"
msg="<font color=FF0000>�i�S�O�޲z��ʡj</font><font color=FFC01F>"&st&"</font>�c�ư��h�Q�P�J�c24�p�ɡC(�����{�D)[����H��="&info(0)&"]"
call boot(st)
end if
if opt="�ʸT" then
conn.execute("update �Τ� set ���A='�ʸT',�n��=now()+8000 where id="&stt)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & st & "','�ʸT',now(),'�S�O�޲z���')"
msg="<font color=FF0000>�i�S�O�޲z��ʡj</font><font color=FFC01F>"&st&"</font>�c�ư��h�Q�a�J�F�Q�K�h�a���C(�L���{�D)[����H��="&info(0)&"]"
call boot(st)
end if
if opt="��ı" then
conn.execute("update �Τ� set ���A='�v',�n��=now()+8000 where id="&stt)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & st & "','�ʸT',now(),'�S�O�޲z���(��ı2024�~')"
msg="<font color=FF0000>�i���v�j</font><font color=FFC01F>"&st&"</font>�c�ư��h�Q�a�J�F�Q�K�h�a���A�}�T��2024�~�C[����H��="&info(0)&"]"
call boot(st)
end if
if opt="���u" then
conn.execute("update �Τ� set ���A='�c',�n��=now()+8888 where id="&stt)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" & st & "','�ʸT',now(),'�S�O�޲z���(�H���]�o')"
msg="<font color=FF0000>�i�H���]�o�j</font>�@�}�q�{�p��A<font color=FFC01F>"&st&"</font>�]�H�Ϧ���W�x�A�Q���X��ѫ�....[����H��="&info(0)&"]"
'conn.execute("delete * from �Τ� where id="&stt)
call zhadan(st)
end if
if opt="�_��" then
conn.execute("update �Τ� set �O�@=false,�_���׽m=0,�_��='�F�C������' where id="&stt)
msg="��������G<a href=javascript:parent.sw('[" & st & "]'); target=f2><font color="& addwordcolor & "><b>" & st & "</b></font></a>�{�b�֦�������_<font color=red>�F�C������</font>,�o�_���O�����`���M�K�j���n�_�A�~�����W�W��쪺�ڡA���C�C�C�U��L�h���۪��ܡI[����H��="&info(0)&"]"
end if
if opt="���y" then
conn.execute("update �Τ� set �s��=�s��+30000000 where id="&stt)
'msg="<font color=FFC01F>�i�������y�j</font><font color=FFC01F>"&st&"</font>��{�u�q�A�|�����\,�������{<font color=FFC01F>"&st&"</font>���a���e�W����^�m���y���I"
end if
if opt="�@��" then
conn.execute("update �Τ� set �s��=�s��-30000000 where id="&stt)
msg="<font color=FFC01F>�i�����@�ڡj</font><font color=FFC01F>"&st&"</font>��{�c�H�A�Y���o�æ���A�v�T���a�M�w�@��3000�U�H���g�١I"
end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing

'��H
sub boot(st)
Application.Lock
onlinelist=Application("ljjh_onlinelist")
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(st) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist")=listnull
else
Application("ljjh_onlinelist")=newonlinelist
end if
Application("ljjh_useronlinename")=useronlinename
Application("ljjh_chatrs")=onliners
Application.UnLock
end sub
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
sd(196)="FFC01F"
sd(197)="FFC01F"
sd(198)="��"
sd(199)="<font color=FFC01F>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.Lock
'���u
sub zhadan(st)
Application.Lock
onlinelist=Application("ljjh_onlinelist")
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(st) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "

if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist")=listnull
else
Application("ljjh_onlinelist")=newonlinelist
end if
Application("ljjh_useronlinename")=useronlinename
Application("ljjh_chatrs")=onliners
Application.UnLock
end sub
Response.redirect "onlinelist1.asp"
%>