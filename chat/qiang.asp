<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
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
if Application("ljjh_qiang")=0 then
	Response.Write "<script Language=Javascript>alert('���ܡG�A�ӱߤF�A��j�Q�H�m���F...');</script>"
	response.end
end if
Application.Lock
Application("ljjh_qiang")=0
Application.UnLock
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from ���~ where ���~�W='����j' and �֦���=" &info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,���O,��O,�ƶq,�Ȩ�,����,sm,�|��) values ('����j',"&info(9)&",'�j��',0,0,1000,0,0,1,2000000,1400,1400,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+1,�|��="&aaao&" where ���~�W='����j' and �֦���="&info(9)
	end if

qiang="["&name&"]�u�O�F�`�A�~�M�q�����⤤�m���o��@�����j,�i���S���l�u�A�n�l�u�h���g��a�γ\����@��!"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
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
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="��"
sd(199)="<font color=9FDF70>�i��������j</font>"&qiang
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
%>
