<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
	response.end
end if
if Application("ljjh_gg")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�p�l�ӱߤF�A�������׳����H����F...');</script>"
	response.end
end if
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
randomize timer
r=int(rnd*3)+1
if r=2 then
rs.open "select id from ���~ where ���~�W='�B�����y',�|�� and �֦���=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,�|��) values ('�B�����y',"&info(9)&",'���~',0,0,0,0,0,1,200000,99008,"&aaao&")"

	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)

	end if
yb="["&info(0)&"]�U��n�ְ�,�@�U�l�N�⪯�������F,�ٱq�����۸̧��@���m�B�����y�n<img src='../hcjs/jhjs/001/99008.gif' border='0'>�j�a�U��n�ְ�!"
else
rs.open "select ����,��O�[ FROM �Τ� WHERE id="&info(9),conn
tlj=(rs("����")*1500+5260)+rs("��O�[")
conn.execute "update �Τ� set ��O=" & tlj & ",�Ȩ�=�Ȩ�+500000  where id="&info(9)
yb="["&info(0)&"]�U��n�ְ�,�@�U�l�N�⪯�������F,���ץͳ�F�@�z�G�b,�����Y�ٽ�F�Ȩ�500000��Ȥl�A�j�a�U��n�ְ�!"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Application.Lock
Application("ljjh_gg")=0
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
sd(199)="<font color=B0D0E0>�i��������j</font>"&yb
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
%>
 