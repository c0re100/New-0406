<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if
to1=request.form("to1")
towho=request.form("to1")
toname=request.form("to1")
if toname="�j�a" or toname=info(0) or toname="" then
	Response.Write "<script language=JavaScript>{alert('�M�k���ﹳ�����A�ЬݥJ�ӤF�I�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
fla=rs("�k�O")
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o1000�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<15 then
rs.close
set rs=nothing	
conn.close
set conn=nothing
ss=15-sj
Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-1000,�ާ@�ɶ�=now() where id="&info(9)
randomize()
r=int(rnd*25)+1
select case r
case 1
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�T����' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�T����',"&info(9)&",'�k��',0,0,1000,2001,2001,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�b"&towho&"�a�̪��ɩ��U����F<font color=red>�k���T����</font>�A�H���p�a�}�]�F."
case 2
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�}���@' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�}���@',"&info(9)&",'�k��',0,0,1000,2002,2002,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0)& "�A���i"&towho&"�a�̧�<font color=red>�k���}���@</font>�q��⤤�m���A�u�O�@�Ƥ�����."
case 3
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='��w�l' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('��w�l',"&info(9)&",'�k��',0,0,1000,2004,2004,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�b����̪��@�ӯ}�q���o�{�F�Q�O�H���<font color=red>�k����w�l</font>."
case 4
rs.close
rs.open "select * FROM ���~ WHERE ���~�W='�m�T�O' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�m�T�O',"&info(9)&",'�k��',0,0,1000,2005,2005,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�u�O���B�A���o�@����<font color=red>�k���m�T�O</font>����Q�A��A�u�O���֤��H��."

case 5
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='���_��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���_��',"&info(9)&",'�k��',0,0,1000,2006,2006,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�o�{"&towho&"�f�U�̬����{�{,����@�N��ӬO�@��<font color=red>���_��</font>,��O�ߤF���H���Y��i"&towho&"���f�U,����_�۵��s���F."
case 6
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='���_��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���_��',"&info(9)&",'�k��',0,0,1000,2007,2007,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�o�{"&towho&"�f�U�̺���{�{,����@�N��ӬO�@��<font color=red>���_��</font>,��O�ߤF���H���Y��i"&towho&"���f�U,����_�۵��s���F."
case 7
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='���_��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('���_��',"&info(9)&",'�k��',0,0,1000,2008,2008,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�o�{"&towho&"�f�U���ť��{�{,����@�N��ӬO�@��<font color=red>���_��</font>,��O�ߤF���H���Y��i"&towho&"���f�U,�����_�۵��s���F."

case 8
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�]�O�p��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�]�O�p��',"&info(9)&",'�k��',0,0,100,1100,1100,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�b����̪��@�ʤd�~�j�𪺾�K�W�o�{�@��<font color=red>�]�O�p��</font>�A"&info(0)&"�A�������u�O�y��."
case 9
rs.close
rs.open "select id FROM �Τ� WHERE �m�W='"&toname&"'"
idd=rs("id")
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�ͤ�J�|' and �֦���=" & idd
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�ͤ�J�|',"&idd&",'�k��',0,0,100,2009,2009,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id 
	end if
xunfaqi=info(0) & "�A�b�h"&towho&"�a�L�ͤ�,�e�F�@�ӳJ�|��"&towho&"!"
case 10
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�����M' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�����M',"&info(9)&",'�k��',0,0,1000,2010,2010,0,0,0,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
xunfaqi=info(0) & "�A�h"&towho&"�a���,�b����o�{<font color=red>�@�⵴���M</font>."
case else
xunfaqi=info(0) & "�A�B��u�O�t�t�r�A���򳣨S���."
end select
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
sd(194)=to1
sd(195)=info(0)
sd(199)="<font color=CFE0B0>�i�M��k���j"& xunfaqi &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>