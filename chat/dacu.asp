<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
response.end
end if
if Application("ljjh_bianfu")=0 then
Response.Write "<script Language=Javascript>alert('���ܡG�A�ӱߤF,�d�����ȤF..');</script>"
response.end
end if
name=lcase(trim(request("name")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize()
r=int(rnd*8)+2
select case r
case 1
mess="<font color=red>�i"& nickname &"�j</font>���_�Ҥl�Q��"&name&"���d���A���G["&name&"]���d���~�ʤj�o�A���R�b�]�A<font color=red>�i"& nickname &"�j</font>�l�F�b�ѳ��S�l�W�A�u�ɤH��,��O����100�A���O�l��200�I"
conn.execute "update �Τ� set ��O=��O-100,���O=���O-200 where id="&info(9)
case 2,3,4,6,7,8,9,10
sql="select id from �d�����A where user='"&name&"'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof then
Response.Write "<script Language=Javascript>alert('"&name&"���d�������F�r�I');</script>"
response.end
else
id=rs("id")
conn.execute "delete * from �d�����A where id="& id &" and user='"&name&"'"
end if
rs.close
rs.open "select id from ���~ where ���~�W='�l�u' and �֦���=" &info(9),conn
If Rs.Bof OR Rs.Eof then
conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�l�u',"&info(9)&",'�u��',0,0,100,2012,0,0,1,2000000,"&aaao&" )"
else
id=rs("id")
conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='�l�u' and �֦���="&info(9)
end if
mess="<font color=red>�i"& nickname &"�j</font>���_�Ҥl���["&name&"]���d���@�}�r�~�A["&name&"]���d���ש�G�s���n��˦b�a�W�A�i"& nickname &"�j�q["&name&"]���d���r��̧��@���l�u!"
case else
mess="<font color=red>�i"& nickname &"�j</font>���_�Ҥl�Q��"&name&"���d���A���G["&name&"]���d���~�ʤj�o�A���R�b�]�A<font color=red>�i"& nickname &"�j</font>�l�F�b�ѳ��S�l�W�A�u�ɤH��,��O����100�A���O�l��200�I"
conn.execute "update �Τ� set ��O=��O-100,���O=���O-200 where id="&info(9)
end select

Application.Lock
Application("ljjh_bianfu")=0
Application.UnLock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application.Lock
Application("ljjh_line")=line+1
Application.UnLock
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="660099"
sd(197)="660099"
sd(198)="��"
sd(199)="<font color=red>�i��������j</font>"&mess
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
