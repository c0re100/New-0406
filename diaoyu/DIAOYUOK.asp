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
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"diaoyu/diao.asp")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
abc=DateDiff("s",Session("diaoyu"),now())
if abc<30 or abc>40 then 
	Response.Write "<script Language=Javascript>alert('�A�Q�@����O�H�o�̤���@�����I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ� from �Τ� where id="&info(9),conn
sj=DateDiff("n",rs("�ާ@�ɶ�"),now())
if sj<1 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('�е��W"&ss&"���A�ӳ����a�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update �Τ� set �ާ@�ɶ�=now()  where id="&info(9)
randomize timer
r=int(rnd*12)+1
'r=9
select case r
case 1
	mess="����"& info(0) &"����@���j�T���A�i�@���t���ϥΡA������O1200�B���O800"
	rs.open "select id from ���~ where ���~�W='�j�T��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j�T��',"&info(9)&",'�t��',0,0,100,300,300,0,-1200,-800,1,20000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=20000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 2
	mess="����"& info(0) &"����@���������A�춰����o5�U�Ȥl"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+50000 where id="&info(9) 
case 3
	mess="����"& info(0) &"����@���j�󳽡A��U�i�H�W�[��O300�B���O50"
	rs.open "select id from ���~ where ���~�W='�j��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j��',"&info(9)&",'�ī~',0,0,100,301,301,0,300,50,1,2000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=2000,�ƶq=�ƶq+1,�|��="&aaao&" where id="&id &" and �֦���="&info(9)
	end if
	rs.close
case 4
	mess="����"& info(0) &"����@���p�U���A��U�i�H�W�[��O100�B���O30"
	rs.open "select id from ���~ where ���~�W='�p�U��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�p�U��',"&info(9)&",'�ī~',0,0,100,302,302,0,100,30,1,500,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=500,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 5
	mess="���I�I"& info(0) &"���S����A�L��e����O���300�A���O���100"
	conn.execute "update �Τ� set ��O=��O-300,���O=���O-100 where id="&info(9)
case 6
	mess= info(0) &"������������Q�D�H�o�{�A�@�}�ޥ��A��O�U��500�B���O�U��200"
	conn.execute "update �Τ� set ��O=��O-500,���O=���O-200 where id="&info(9)
case 7
	mess="����"& info(0) &"�B��u�O�n��BiangBiang�n�r�I����j�T���B�j�󳽡B�p�U���U�@���I"

	rs.open "select id from ���~ where ���~�W='�j�T��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j�T��',"&info(9)&",'�t��',0,0,100,300,300,0,-1200,-800,1,20000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=20000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close

	rs.open "select id from ���~ where ���~�W='�j��' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j��',"&info(9)&",'�ī~',0,0,100,301,301,0,300,50,1,2000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=2000,�ƶq=�ƶq+1,�|��="&aaao&" where id="&id &" and �֦���="&info(9)
	end if
	rs.close

	rs.open "select id from ���~ where ���~�W='�p�U��',�|��="&aaao&"  and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�p�U��',"&info(9)&",'�ī~',0,0,100,302,302,0,100,30,1,500,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=500,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 8

   mess="�A�b�e��ߨ�@���}�ª��m�ƶ����y�n<img src='../hcjs/jhjs/001/99004.gif' border='0'>�A"&info(0)&"�A�u�O���B��."
rs.open "select id from ���~ where ���~�W='�ƶ����y' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��) values ('�ƶ����y',"&info(9)&",'���~',0,0,0,0,0,1,200000,99004,99004,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if

rs.close
case 9
	mess="����"& info(0) &"�b����ɵo�{�F�@�Ӭy�P�A"& info(0) &"�}�߷��F!"
	rs.open "select id from ���~ where ���~�W='�y�P' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�y�P',"&info(9)&",'���~',0,0,1000,111111,111111,0,0,0,1,1000000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=1000000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 10
	mess="����"& info(0) &"�b�M�䱼�i���̪��Ȩ�ɵo�{�F�@�Ӭy�P�\�A�u�O���B!"
	rs.open "select id from ���~ where ���~�W='�y�P�\' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�y�P�\',"&info(9)&",'���~',0,0,1000,2003,2003,0,0,0,1,1000000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=1000000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 11
	mess="����"& info(0) &"�b�e��o�{�F�@���s�]�A�u�O���!"
	rs.open "select id from ���~ where ���~�W='�s�]' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�s�]',"&info(9)&",'���~',0,0,100,50000,50000,0,0,0,1,1000000,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �Ȩ�=1000000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 12
	mess="����"& info(0) &"�b�������ɭԡA����H�e�F10�U�|�O���L�A"& info(0) &"�èS�����x�A�Ϧӧ⨺10�U�|����i�ۤv���f�U�̡C"
	conn.execute "update �Τ� set �|���O=�|���O+100000 where id="&info(9) 
case 13
	mess="����"& info(0) &"����@�Ӳ��l�A���̥����O�����A�g�ƺ��A�����`�Ƴ���10�U�ӡA"& info(0) &"�o�@�ͥi�H�a�o�Ǫ��P�L���A�u�O�O�H�r�}�C"
	conn.execute "update �Τ� set ����=����+100000 where id="&info(9) 
case else
   mess="�A�~�ۨS�Ʃ��e����A���ƳQ�A�b�e���o�{�@���S�ت���<img src='../hcjs/jhjs/001/2003303.gif' border='0'>�A"&info(0)&"�A���鲴�n�y��!�K�K!"
rs.open "select id from ���~ where ���~�W='�S�ت���' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��) values ('�S�ت���',"&info(9)&",'���~',0,0,0,0,0,1,10000,2003303,2003303,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if

rs.close
end select


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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=#ff0000>����</font>"& info(0) &"�b�e�䳨���G"&mess
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>�������\!</title></head>

<body oncontextmenu=self.event.returnValue=false background="../jhimg/bk_hc3w.gif">
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