<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.B0D0E0irect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
nickname=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
	response.end
end if
if Application("ljjh_bianfu")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�p�l�ӱߤF�A�ǳ����Q�H�����F...');</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize()
r=int(rnd*12)+2
select case r
case 1
sql="select * from ���~ where ���~�W='�B���Ȱw' and �֦���="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('�B���Ȱw',"&info(9)&",'�t��',0,0,2000,-2000,-3000,1,20000,1300,"&aaao&")"
	else
		
		conn.execute "update ���~ set �Ȩ�=20000,�ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='�B���Ȱw' and �֦���="&info(9)
	end if
mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��@��B���Ȱw�A�i�@���t���ϥΡA������O3000�B���O2000"
case 2
	mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��@�]�����A�춰����o�Q�U�Ȥl"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+100000 where id="&info(9) 
case 3
	mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��@���Q�������A�ܽ�o��Ȩ�60000"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+60000 where id="&info(9)
case 4
	mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��@��ï]�A�ܽ�o��Ȥl4000��"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+4000 where id="&info(9)
case 5
	mess="�˾`��<font color=B0D0E0>�i"& nickname &"�j</font>�ǳ��S�����A�ӥB���p�߽�쳴��,��O���500�A���O���200"
	conn.execute "update �Τ� set ��O=��O-500,���O=���O-200 where id="&info(9)
case 6
	mess="�ǳ��s�F����,<font color=B0D0E0>�i"& nickname &"�j</font>�ϧܾD��r��,��O�U��1000�B���O�U��500"
        conn.execute "update �Τ� set ��O=��O-1000,���O=���O-500 where id="&info(9)
case 7
        mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�B��u�O�n�����o�F�r�I�����ǳ���o��F�@�j�������,������1000000��Ȥl�I"
        conn.execute "update �Τ� set �Ȩ�=�Ȩ�+1000000 where id="&info(9)
case 11
       mess="<font color=B0D0E0>�i"& nickname &"�j</font>�B��u�O�n�A�����ǳ���o��F<font color=B0D0E0>������_�F�C������</font>�I�j�a�٤��ַm�I"
      conn.execute "update �Τ� set �_��='�F�C������',�_���׽m=0 where id="&info(9)
case 9

sql="select id from ���~ where ���~�W='�^�R��' and �֦���="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('�^�R��',"&info(9)&",'�ī~',0,0,2000,0,100000,1,200000,509,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��}�@�_��[�^�R��]�A�A�Υi�W�[��O100000"
case 10
sql="select id from ���~ where ���~�W='�Ѥs����' and �֦���="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('�Ѥs����',"&info(9)&",'�ī~',0,0,3000,100000,0,1,200000,500,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&"  where id="& id &" and �֦���="&info(9)
	end if
mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��}�@�_��[�Ѥs����]�A�A�Υi�W�[���O100000"
case 11
sql="select id from ���~ where ���~�W='�r����' and �֦���="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,��T��,����,����,���s,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('�r����',"&info(9)&",'�r��',0,0,4000,0,10000,1,200000,1200,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&"  where id="& id &" and �֦���="&info(9)
	end if
mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��u���r��[�r����]�A�i���˹����O10000"
case 12
sql="select id from ���~ where ���~�W='���s���y' and �֦���="&info(9)
Set Rs=conn.Execute(sql)
	
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,��T��,����,����,���s,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('���s���y',"&info(9)&",'���~',0,0,4000,0,10000,1,200000,99007,"&aaao&")"
	else
id=rs("id")	
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&"  where id="& id &" and �֦���="&info(9)
	end if
mess="����<font color=B0D0E0>�i"& nickname &"�j</font>�����ǳ���o��@���m���s���y�n<img src='../hcjs/jhjs/001/99007.gif' border='0'>�A�����ЫB½�������ۡI"
case else
 conn.execute "update �Τ� set �_��='�F�C������',�_���׽m=0 where id="&info(9)
 mess="<font color=B0D0E0>�i"& nickname &"�j</font>�B��u�O�n�A�����ǳ���o��F<font color=B0D0E0>������_�F�C������</font>�I�j�a�٤��ַm�I"
       
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
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i��������j</font>"&mess
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
