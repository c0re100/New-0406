<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ��O,���A from �Τ� where id="&info(9),conn
if rs("��O")<0 or rs("���A")="��" then Response.Redirect "../exit.asp"
sja=session("wabao")
randomize timer
r=int(rnd*12)+8
  if session("yuokaa")="" then
     Response.Redirect "diao.asp"
     response.end                                                           
  end if
session("wabao")=""
session("yuokaa")=""
select case r

case 1
	mess="����"&info(0)&"����@�⭸��C�A�i�@���t���ϥΡA������O80000�B���O10000"
rs.close
rs.open "select id from ���~ where ���~�W='����C' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,����,sm,��T��,����,���s,��O,���O,�ƶq,�|��) values ('����C',"&info(9)&",'�t��',1000,1000,100,0,0,80000,10000,1,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
				
		        end if
case 2
	mess="����"&info(0)&"����@��ɵo¯�A�i�H�ΰ��˹��~�A�[���s1000�B����1000"
rs.close
rs.open "select id from ���~ where ���~�W='�ɵo¯' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,����,��T��,���O,��O,���s,�ƶq,�Ȩ�,����,sm,����,�|��) values ('�ɵo¯',"&info(9)&",'�˹�',1000,100,0,0,1000,1,200000,1500,1500,45,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
				
		        end if
case 3
	mess="����"&info(0)&"����@��<img src='../hcjs/jhjs/001/135.gif' border=0 >�d�~�H�ѡA�A�Υi�W�[���O100000�I"
rs.close
rs.open "select id from ���~ where ���~�W='�d�~�H��' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,��T��,���O,��O,����,���s,����,sm,�ƶq,�Ȩ�,�|��) values ('�d�~�H��',"&info(9)&",'�ī~',2000,100000,0,0,0,135,135,1,2000000,"&aaao&")"
			
                        else
id=rs("id") 
				conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
				
		        end if
case 4
	mess="����"&info(0)&"����@��<img src='../hcjs/jhjs/001/136.gif' border=0 >�U�~�F�ۡA�A�Υi�W�[��O200000�I"
rs.close
rs.open "select id from ���~ where ���~�W='�U�~�F��' and �֦���=" & info(9),conn
			if rs.eof and rs.bof then
			conn.execute "insert into ���~(���~�W,�֦���,����,���O,��O,����,���s,��T��,����,sm,�ƶq,�Ȩ�,�|��) values ('�U�~�F��',"&info(9)&",'�ī~',0,200000,0,0,100,136,136,1,5000000,"&aaao&")"
			
                        else 
id=rs("id")
				conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
				
		        end if
case 5
rs.close
rs.open "select id from ���~ where ���~�W='�]�O�p��' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,�|��) values ('�]�O�p��',"&info(9)&",'�k��',5000,0,0,0,0,1,200000,1100,1100,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if
   mess="�A�b�@�����Y�U������@���]�O�p�ۡA"&info(0)&"�A�������u�O�y��."
case 6
	mess="����"&info(0)&"����@����_�͡A�ܽ�o��Ȩ�50000"

   conn.execute("update �Τ� set �Ȩ�=�Ȩ�+50000 where id="&info(9))

case 7
	mess="����"&info(0)&"����@�����H���Y�A�𦺤F�A�D�w�U��20"
    conn.execute("update �Τ� set �D�w=�D�w-20 where id="&info(9))

case 8
	mess="�˾`��"&info(0)&"�_�èS���A�ӥB���p�߽�쳴��,���O���200"
   conn.execute("update �Τ� set ���O=���O-200 where id="&info(9))

case 9

mess="����"&info(0)&"����@�]�����A�춰����o�Q�U�Ȥl"

    conn.execute "update �Τ� set �Ȩ�=�Ȩ�+100000 where id="&info(9)
case 10
        mess=""&info(0)&"�B��u�O�n�A�~�M���X�F������_�F�C�����ۡI"
        conn.execute "update �Τ� set �_��='�F�C������',�_���׽m=0 where id="&info(9)
case 11
rs.close
rs.open "select id from ���~ where ���~�W='�������y' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,�|��) values ('�������y',"&info(9)&",'���~',0,0,0,0,0,1,200000,99001,99001,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if
   mess="�A�b�@�����Y�U������@���}�ª��m�������y�n<img src='../hcjs/jhjs/001/99001.gif' border='0'>�A"&info(0)&"�A�u�O���B��."
case 12
rs.close
rs.open "select id from ���~ where ���~�W='����' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,�|��) values ('����',"&info(9)&",'���~',0,0,0,0,0,1,10000,2003307,2003307,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if
   mess="�A�b�@�����Y�U������@������<img src='../hcjs/jhjs/001/2003307.gif' border='0'>�A"&info(0)&"�A�٤��֨��X�ӧr,�����i�H�Ψӷڳy���@�Z���r,�֦��æn�I"
case else
    mess="�j�s�S�ӷm�T,"&info(0)&"�ϧܾD��r��,�Ĩĥ�X1000��"
   conn.execute("update �Τ� set �s��=�s��-1000 where id="&info(9))
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="��"
sd(199)="<font color=red>�i�����j</font><font color=FFCFCF>"&info(0)&"</font>�b��s���_�áG<font color=FFCFCF>"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>���_��OK</title></head>

<body oncontextmenu=self.event.returnValue=false bgcolor="#000000">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=�������f onClick='window.close()' name="button">
</div>
</td>
</tr>
</table>
</div>
</body>
</html>
