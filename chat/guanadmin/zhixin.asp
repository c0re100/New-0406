<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(2)<8  then Response.Redirect "../../error.asp?id=425"
pass1=Request.Form("pass1")
pass11=Request.Form("pass11")
ljgl=Request.Form("ljgl")
if InStr(pass1,"'")<>0 or InStr(pass1,"`")<>0 or InStr(pass1,"=")<>0 or InStr(pass1,"-")<>0 or InStr(pass1,",")<>0 or InStr(pass1,"\")<>0 or InStr(pass1,"/")<>0 or InStr(pass1,"|")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');history.back();}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT grade FROM �Τ� where �m�W='"&pass1&"' and ���A='���`'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('��p�A�ҭn�쪺�H�b�ƾڮw�̨S���I�Ьd�ݦW�r�O�_���T�I');history.back();</script>"
Response.End
else
 if rs("grade")>info(2) then
	
 Response.Write "<script language=javascript>alert('��p�A�ҭn�쪺�H�޲z�Ť�A���A�A�줣�F���I');history.back();</script>"
Response.End
 end if 
select case ljgl
case "�e��"
sql="update �Τ� set ���A='��',�n��=now() where �m�W='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" &pass1& "','�e��',now(),'" &pass11& "')"
msg="<font color=FF0000>�i�S�O�޲z��ʡj</font>�ѪŤ���M���X�@������A�⤣ť�ܪ��p��<font color=FFFDAF>"&pass1&"</font>��J�F�x���j�c�C(�e��)�]��]�G<font color=FFFDAF>"&pass11&"</font>�^[����H��="&info(0)&"]"
call boot(pass1)
case "���c"
sql="update �Τ� set ���A='�c',�n��=now()+1/84 where �m�W='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" &pass1& "','���c',now(),'" &pass11& "')"
msg="<font color=FF0000>�i�S�O�޲z��ʡj</font>�H�ۤ@�n�ªZ���[�ܡA�x�����H�p�T���ꪺ�N��ť�ܪ��p��<font color=FFFDAF>"&pass1&"</font>���i�F���Ǥj�c�C(���c)�]��]�G<font color=FFFDAF>"&pass11&"</font>�^[����H��="&info(0)&"]"
call boot(pass1)
case "�ʸT"
sql="update �Τ� set ���A='�ʸT',�n��=now()+8000 where �m�W='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" &pass1& "','�ʸT',now(),'" &pass11& "')"
msg="<font color=FF0000>�i�S�O�޲z��ʡj</font><font color=FFFDAF>"&pass1&"</font>�c�ư��h�Q�a�J�F�Q�K�h�a���C(�ʸT)�]��]�G<font color=FFFDAF>"&pass11&"</font>�^[����H��="&info(0)&"]"
call boot(pass1)
'case "��ı"
'sql="update �Τ� set ���A='�v',�n��=now()+8000  where �m�W='"&pass1&"'"
'Set Rs=conn.Execute(sql)
'conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" &pass1& "','�ʸT',now(),'" &pass11& "')"
'msg="<font color=FF0000>�i�S�O�޲z��ʡj</font><font color=FFFDAF>"&st&"</font>�c�ư��h�Q�a�J�F�Q�K�h�a���A�}�T��2024�~�C[����H��="&info(0)&"]"
'call boot(pass1)
case "�_��" 
sql="update �Τ� set �O�@=false,�_���׽m=0,�_��='�F�C������' where �m�W='"&pass1&"'"
Set Rs=conn.Execute(sql)
msg="��������G<a href=javascript:parent.sw('[" & pass1 & "]');target=f2><font color=FFFDAF><b>" & pass1 & "</b></font></a>�{�b�֦�������_<font color=FFFDAF>�F�C������</font>,�o�_���O�����`���M�K�j���n�_�A�~�����W�W��쪺�ڡA���C�C�C�U��L�h���۪��ܡI"
'conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" &pass1& "','���_��',now(),'" &pass11& "')"
case "���u"
sql="update �Τ� set ���A='�c',�n��=now()+8888 where �m�W='"&pass1&"'"
Set Rs=conn.Execute(sql)
conn.execute "insert into �ư�(�޲z����,�޲z�ﹳ,�޲z�覡,�欰�ɶ�,��]) values ('" & info(0) & "','" &pass1& "','�ʸT',now(),'���u:" &pass11& "')"
msg="<font color=FF0000>�i�H���]�o�j</font>�@�}�q�{�p��A<font color=FFFDAF>"&pass1&"</font>�]�H�Ϧ���W�x�A�Q���X��ѫ�....�]��]�G<font color=FFFDAF>"&pass11&"</font>�^[����H��="&info(0)&"]"
Application.Lock
Application("ljjh_bombname")=Application("ljjh_bombname") & " " & pass1 & " "
Application.UnLock
call zhadan(pass1)
end select

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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.Lock
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script language=javascript>alert('["&ljgl&"]�G["&pass1&"]����ާ@���\�I');history.back();</script>"
end if
'��H
sub boot(pass1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(pass1) then
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
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
end sub
'���u
sub zhadan(pass1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(pass1) then
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
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
end sub
%>