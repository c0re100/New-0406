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
if Application("ljjh_dalie")<>"�Ѫ�" then
	Response.Write "<script Language=Javascript>alert('���ܡG�{�b�٨S���y���i�H���A�Ф@�|�A�ӧa�C');window.close();</script>"
	response.end
end if
Application.Lock
Application("ljjh_dalie")="�L"
Application.UnLock
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "update �Τ� set ���O=���O+800,��O=��O+1200,allvalue=allvalue+1,�Ȩ�=�Ȩ�+5000 where id="&info(9)
rs.open "select id from ���~ where ���~�W='�Ѫ��' and �֦���=" &info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��) values ('�Ѫ��',"&info(9)&",'�ī~',0,0,0,150,150,1,2000,400,400,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update ���~ set �Ȩ�=2000,�ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='�Ѫ��' and id="&id
end if
mess="<b><font color=red>���~�w�g�Q["&info(0)&"]�����A�u�O�^���֦~�r�A�Фj�a��߲�ѧa!</b>"
randomize timer
r=int(rnd*3)+1
if r=3 then 
rs.close
rs.open "select id from ���~ where ���~�W='���믵�Z' and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��) values ('���믵�Z',"&info(9)&",'���~',0,0,0,0,0,1,200000,99003,99003,0,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if
mess="<b><font color=red>���~�w�g�Q["&info(0)&"]�����A�u�O�^���֦~�r�A�Фj�a��߲�ѧa,"&info(0)&"�b���~�ۧ��@��<font color=FFFDAF>�m���믵�Z�n<img src='../hcjs/jhjs/001/99003.gif' border='0'></font></font></b>"
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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>�i�n�����j"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<title>���y���\</title>
<style type="text/css">
<!--
table {  border: #000000; border-style: solid; border-top-width: 1px; border-right-width: 1px; border-bottom-width: 1px; border-left-width: 1px}
font {  font-size: 12px}
-->
</style>

<body text="#000000" background="../../images/8.jpg"" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<table width="98%" border="0" height="156" bordercolor="#330033" cellspacing="0" cellpadding="0" align="center">
<tr>
<td height="17" bgcolor="#0000FF" align="center"><font color="#FFFFFF">���y���\</font></td>
</tr>
<tr>
<td bgcolor="#FFCC99" align="center" height="157"><font> �z���\���N���~�����A�W�[���O+800�A��O+1200�A�Ȩ�+5000�A�I��+1�C<br>
<br>
��פ@���A�ϥΪ��~:���O+150,��O+150</font></td>
</tr>
</table>
<p>&nbsp;</p>
</body>
</html>
