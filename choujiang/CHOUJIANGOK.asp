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
if InStr(http,"choujiang/chou.asp")=0 then 
Response.Write "<script language=javascript>{alert('�藍�_�A�{�ǩڵ��z���ާ@�I�I�I\n     ���T�w��^�I');parent.history.go(-1);}</script>" 
Response.End 
end if
abc=DateDiff("s",Session("choujiang"),now())
if abc<30 or abc>240 then 
	Response.Write "<script Language=Javascript>alert('�w�g�L�F����ɶ��A�ЫݤU���a�I');location.href = 'javascript:history.go(-1)';</script>"
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
	Response.Write "<script Language=Javascript>alert('�е��W"&ss&"���A�ӽ�B��a�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update �Τ� set �ާ@�ɶ�=now()  where id="&info(9)
randomize timer
r=int(rnd*12)+1
'r=9
select case r
case 1
	mess="��o�F�j���A���~�O�����B���s�B��O�U�[3��"
	conn.execute "update �Τ� set ����=����+300000000,���s=���s+300000000,��O=��O+300000000 where id="&info(9)

case 2
	mess="��o�F�G���A���~�O�g��5000��"
	conn.execute "update �Τ� set allvalue=allvalue+500000000000 where id="&info(9) 
case 3
	mess="��o�F�T���A���~�O�k�O10��"
	conn.execute "update �Τ� set �k�O=�k�O+1000000000 where id="&info(9)
case 4
	mess="��o�F�|���A���~�O�D�w1��"
	conn.execute "update �Τ� set �D�w=�D�w+100000000 where id="&info(9)
case 5
	mess="��o�F�����A���~�O���I10�U"
	conn.execute "update �Τ� set ���I=���I+100000 where id="&info(9)
case 6
	mess="��o�F�����A���~�O�Ȩ�10��"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+1000000000 where id="&info(9)
case 7
	mess="��o�F�C���A���~�O�Ѳ�"
	conn.execute "update �Τ� set �Ѳ�=1 where id="&info(9)
case 8

        mess="��o�F�K���A���~�O��W"
	conn.execute "update �Τ� set �ʧO='�L' where id="&info(9)
case 9
	mess="��o�F�E���A���~�O�O�s�M10��"
	rs.open "select id from ���~ where ���~�W='�O�s�M' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,sm,�|��) values ('�O�s�M',"&idd&",'�k��',0,0,0,3400000,900000,20032017,10,300,20032017,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+10,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 10
	mess="��o�F�Q���A���~�O�ܴL�C10��"
	rs.open "select id from ���~ where ���~�W='�ܴL�C' and �֦���="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,���s,���O,��O,��T��,����,����,�ƶq,����,sm,�|��) values('�ܴL�C',"&idd&",'����',800000,0,0,3200000,0,20032016,10,265,20032016,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+10,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
	rs.close
case 11
	mess="��o�F�S�O���A���~�O�|�O�M�����U1��"
	conn.execute "update �Τ� set �|���O=�|���O+100000000,����=����+100000000 where id="&info(9)
case else
   mess="��o�F�w�����A���~�O�y�P1000��"
rs.open "select id from ���~ where ���~�W='�y�P' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�y�P',"&info(9)&",'���~',0,0,1000,111111,111111,0,0,0,1000,1000000,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1000,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

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
sd(199)="<font color=#ff0000>����</font>"& info(0) &"�b������ʤ�"&mess
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>���G���G</title></head>

<body oncontextmenu=self.event.returnValue=false background="../jhimg/bk_hc3w.gif" link="#000000" vlink="#000000" alink="#000000">
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
</tr>
</table>
</div>
</body>
<noembed>
</html>