<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if
to1=request.form("to1")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
rs.open "select �k�O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
fla=rs("�k�O")
if fla<100 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o100�I�ڡI');location.href = 'javascript:history.go(-1)';}</script>"
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
conn.execute "update �Τ� set �k�O=�k�O-100,�ާ@�ɶ�=now() where id="&info(9)
randomize()
rnd1=int(rnd()*5)
if rnd1<3 then
banbianshu=info(0) & "�n�i���A�A�M�M�F����U�Ө����]�S����줰������y,"&info(0) & "�l��100�I�k�O!" 
else
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�����y' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,��T��,����,���s,���O,��O,�ƶq,�Ȩ�,����,sm,����,�|��) values ('�����y',"&info(9)&",'�k�_',100,0,0,0,0,1,200000,1600,1600,0,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu=info(0)& "�A�d���U�W�M�X<font color=red>�����y</font>���U���A�o���Ʀb�@�������̳Q�A���F�A����~�~�����y���������]�O�a." 
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
sd(194)=info(9)
sd(195)=info(0)
sd(199)="<font color=CFE0B0>�i�M������j"& banbianshu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>