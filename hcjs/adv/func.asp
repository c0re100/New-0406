<%
function huayuan(name1,sh1,sy1)
%> <!--#include file="data.asp"--> <%
sql="SELECT �_�� FROM �Τ� WHERE �m�W='" & name1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
      huayuan=""
else
      if rs("�_��")="�F�C������" then
            huayuan="" & name1 & "�A�w�g���F�C�����ۤF�����׷ҥH���ׯ��t�q�M�_�I" 
      else
            randomize timer
            rand=Int(79 * Rnd + 10)            
                 if rand>=10 and rand<=50 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�O�H�d�U��100��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+100,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=51 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�O�H�d�U��5000��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+5000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=52 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�O�H�d�U��500��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+500,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=53 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A���F���~�A�O�Ԥ��U�A�����F���~�A���o�ˤF3000���O�I"
                        sql="update �Τ� set ��O=��O-11,���O=���O-3000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=54 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�Q�}�U���Y�@�պL�F�@���A��F100��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�-100,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=55 or rand=85 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�@�_�S���s���ġA�@�𤧤U�A�����@�r�Ӧ��I"                         
                        sql="update �Τ� set ���A='��',�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into �H�R(����,�ɶ�,����,���]) values ('" & name1 & "',now(),'�L','�b�t�q��������r�����Ӧ�')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=56 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�j�N�_�áA���Q�h�����ڡA�_�ê��������}�A�A�Q�ýb�g�ˡA�l����O10000�I"
                        sql="update �Τ� set ��O=��O-10000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=57 or rand=86 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�����^���a�V�A�q������֤F�@�ӵL�W�j�L�I"
                        sql="update �Τ� set ���A='��' where �m�W='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into �H�R(����,�ɶ�,����,���]) values ('" & name1 & "',now(),'�L','�b�t�q�����^���a�V')"
                        conn.execute sql
                      call boot(name1)
                 elseif rand=58 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A���J�@����k�l�Q�Ѫ�l���A�A�^���Ϭ��A�y�O�W��1000�I"
                        sql="update �Τ� set ��O=��O-100,�y�O=�y�O+1000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=59 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A���J�@�@�~���H�A�g���I�A�A�D�w�W��200�I"
                        sql="update �Τ� set ��O=��O-11,�D�w=�D�w+200,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=60 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�~�F�~�y�I�y�O�W��50"
                        sql="update �Τ� set ��O=��O-10,�y�O=�y�O+50 where,�ާ@�ɶ�=now() �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=61 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A���J�@����k�l�A�A��o�ϥX�t�q�A�D�w�W��500�I"
                        sql="update �Τ� set ��O=��O-1,�D�w=�D�w+500,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=62 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�o�{�F�@�Ӥj�_�áA�g�L�@�f�_�I�A�A���\���줭�Q�U�_�äΪZ�\����,���O�W��1000�I"
                        sql="update �Τ� set ���O=���O+1000,�Ȩ�=�Ȩ�+500000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand>=63 or rand<=66 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�O�H�d�U��1000��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+1000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=67 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�~�F�~�y�I�y�O�W��100"
                        sql="update �Τ� set ��O=��O-10,�y�O=�y�O+100,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=68 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�j�_�áA���b�����~���Q�ýb�g���I"
                        sql="update �Τ� set ���A='��' where �m�W='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into �H�R(����,�ɶ�,����,���]) values ('" & name1 & "',now(),'�L','�t�q�����_�îɳQ�ýb�g��')"
                        conn.execute sql
                     call boot(name1)
                 elseif rand=69 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A���r�D�A�����Z�\���j�A����@���I"
                        sql="update �Τ� set ��O=��O-11 where,�ާ@�ɶ�=now() �m�W='" & name1 & "'"
conn.execute sql
                 elseif rand=70 or rand=87 then
huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�J�줳�ĭ襩�]�b���a�A�b�j��30�^�X��A�A���ĦӦ��󤳼ļC�U�I"
                        sql="update �Τ� set ���A='��' where �m�W='" & name1 & "'"
                        conn.execute sql 
                       
                        sql="insert into �H�R(����,�ɶ�,����,���]) values ('" & name1 & "',now(),'�L','�b�t�q�_�I���Q��')"
                        conn.execute sql
                       call boot(name1)
                 elseif rand=71 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�O�H�d�U��1�U��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+10000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=72 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A��M�b�𨤵o�{1��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+1,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=73 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�@�ة_�S���s���ġA�@�𤧤U�A���O�W�i2000�I"
                        sql="update �Τ� set ��O=��O-11,���O=���O+2000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=74 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�O�H�d�U��300��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+300,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=75 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�O�H�d�U��200��Ȥl�I"
                        sql="update �Τ� set ��O=��O-11,�Ȩ�=�Ȩ�+200,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=76 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�~�F�~�y�I�y�O�W��100"
                        sql="update �Τ� set ��O=��O-10,�y�O=�y�O+100,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=77 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�����^�U�s�Y�I�j�������A��O�l��5000�A���j�������A������֡A�o�{5000��Ȥl"
                        sql="update �Τ� set ��O=��O-5000,�Ȩ�=�Ȩ�+5000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=78 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A���J�@�@�~���H�A�g���I�A�A�D�w�W��2000�I"
                        sql="update �Τ� set ��O=��O-11,�D�w=�D�w+2000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=79 or rand=88 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�b�_�I������l�A�O�ԤU�٬O�G�����l�f�U�I"
                        sql="update �Τ� set ���A='��' where �m�W='" & name1 & "'"
                        conn.execute sql 
                                              sql="insert into �H�R(����,�ɶ�,����,���]) values ('" & name1 & "',now(),'�L','�b�t�q�_�I�����`')"
                        conn.execute sql
                        call boot(name1)
                 elseif rand=80 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�o�{�F�W�������O�ߪk�A���O�j�i1000�I�I"
                        sql="update �Τ� set ��O=��O-11,���O=���O+1000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=81 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A���J�@����k�l�A�A��ߤj�_�A���I�ɡA�D�w�U��1000�I"
                        sql="update �Τ� set ��O=��O-1000,�D�w=�D�w-1000,�ާ@�ɶ�=now() where �m�W='" & name1 & "'"
                        conn.execute sql
                 elseif rand=82 or rand=89 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�A�]�䤣�쭹���A�����j���F�A�q������֤F�@�ӵL�W�j�L�I"
                        sql="update �Τ� set ���A='��' where �m�W='" & name1 & "'"
                        conn.execute sql 
                        sql="insert into �H�R(����,�ɶ�,����,���]) values ('" & name1 & "',now(),'�L','�b�t�q�_�I�����`')"
                        conn.execute sql
                     call boot(name1)
                 elseif rand=83 then
                        huayuan="" & name1 & "�b�t�q�W" & sy1 & sh1 & "�A�N�~�a�o�{�F���H�b�t�q�d�U���_���è��a�A���O�̭����򳣨S���I"
             end if    
       end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
end function
sub boot(to1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
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
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
%>