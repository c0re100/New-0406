<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
abc=DateDiff("s",Session("diaoyu"),now())
if abc<90 or abc>115 then 
	Response.Write "<script Language=Javascript>alert('�A�Q�@����O�H�o�̤���@�����I');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ��O,���A,�ާ@�ɶ� from �Τ� where id="&info(9),conn
if rs("��O")<0 or rs("���A")="��" then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Redirect "../exit.asp"
end if
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
conn.execute "update �Τ� set �ާ@�ɶ�=now()  where id="&info(9)
randomize timer
r=int(rnd*6)+1
select case r
case 2
	mess="����"& info(0) &"����������o��@�]�����A�춰����o�Q�U�Ȥl"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+100000 where id="&info(9) 
case 3
	mess="����"& info(0) &"����������o��@���Q�������A�ܽ�o��Ȩ�60000"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+60000 where id="&info(9)
case 4
	mess="����"& info(0) &"����������o��@��ï]�A�ܽ�o��Ȥl4000��"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+4000 where id="&info(9)
case 5
	mess="�˾`��"& info(0) &"�����S�����A�ӥB���p�߽�쳴��,��O���500�A���O���200"
	conn.execute "update �Τ� set ��O=��O-500,���O=���O-200 where id="&info(9)
case 6
	mess="�����s�F����,"& info(0) &"�ϧܾD��r��,��O�U��1000�B���O�U��500"
        conn.execute "update �Τ� set ��O=��O-1000,���O=���O-500 where id="&info(9)
case 7
        mess="����"& info(0) &"�A�B��u�O�n�����o�F�r�I����������o��F�@�j�������,������150000��Ȥl�I"
        conn.execute "update �Τ� set �Ȩ�=�Ȩ�+150000 where id="&info(9)
case 8
        mess=""& info(0) &"�B��u�O�n�A����������~�M���X�F������_�F�C�����ۡI�j�a�٤��ַm�I"
          conn.execute "update �Τ� set �_��='�F�C������',�_���׽m=0 where id="&info(9)
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
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="��"
sd(199)="<font color=FFCFCF>����</font>"& info(0) &"�b�����������G"&mess
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>������OK</title></head>

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