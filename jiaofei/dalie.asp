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
	Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
	response.end
end if
if session("dalie")=false then
	Response.Write "<script Language=Javascript>alert('���ܡG�Q�@���ܡH�I');window.close();</script>"
	response.end
end if
if Application("ljjh_jiaofei")<>"�g��" then
	Response.Write "<script Language=Javascript>alert('���ܡG�{�b�٨S���g��C');window.close();</script>"
	response.end
end if
Application.Lock
Application("ljjh_jiaofei")="�L"
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update �Τ� set ���O=���O+8000,��O=��O+9000,allvalue=allvalue+2,�Ȩ�=�Ȩ�+500000 where id="&info(9)
rs.open "select id from ���~ where ���~�W='�l�u' and �֦���=" &info(9),conn
If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�l�u',"&info(9)&",'�u��',0,0,100,2012,2012,0,0,0,1,2000000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='�l�u' and �֦���="&info(9)
	end if
mess="<b><font color=red>�g��w�g�Q["&info(0)&"]�ĤO�i���A�u�O�^���֦~�r�A�����Գ���o�{�j�O�l�u!</b>"
randomize timer
r=int(rnd*3)+1
if r=3 then 
rs.close
rs.open "select id from ���~ where ���~�W='�B�v�C��' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��) values ('�B�v�C��',"&info(9)&",'���~',0,0,0,0,0,1,200000,99005,99005,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&"  where id="& id &"and �֦���="&info(9)

	end if
mess="<b><font color=red>�g��w�g�Q["&info(0)&"]�ĤO�i���A�u�O�^���֦~�r�A�����Գ���o�{�j�O�l�u�@�o�ñq�g��۱o��@���m�B�v�C�Сn<img src='../hcjs/jhjs/001/99005.gif' border='0'></font></b>"
end if
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�����j"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<title>�ϭꦨ�\</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
-->
</style>

<body text="#000000" background="../images/8.jpg"" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="98%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#0000FF" align="center"><font color="#FFFFFF">�ϭꦨ�\</font></td>
</tr>
<tr>
<td bgcolor="#FFCC99" align="center" height="157"><font> �z���\���N�g�ꥴ���A�W�[���O+8000�A��O+9000�A�Ȩ�+500000�A�԰��g��+2�C<br>
<br>
�����Գ���o�{�j�O�l�u�@�o�A�t�X��j�ϥΡA���ˤO100000</font></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>
