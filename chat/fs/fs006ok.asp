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
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�ҧ������H���b��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
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
if sj<9 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=9-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-100,�ާ@�ɶ�=now() where id="&info(9)
randomize()
'rnd1=int(rnd()*3)
r=int(rnd*7)+4
select case r
case 1
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�B��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�B��',"&info(9)&",'�ī~',0,0,100,151,151,0,150,150,1,2000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A�I�i�k�N<bgsound src=wav/Phant008.wav loop=1>�ܥX�F�@�ӦB���A���Ӫk�O100."
case 2
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�q��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�q��',"&info(9)&",'���~',0,0,100,2014,2014,0,0,0,1,800,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A�I�i�k�N<bgsound src=wav/Phant008.wav loop=1>�ܥX�F�@���q�ۡA���Ӫk�O100."
case 3
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�j�T��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j�T��',"&info(9)&",'�t��',0,0,100,300,300,0,-1200,-800,1,20000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A�I�i�k�N<bgsound src=wav/Phant008.wav loop=1>�ܥX�F�@�Ӥj�T���A���Ӫk�O100."
case 4
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�j��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�j��',"&info(9)&",'�ī~',0,0,100,301,301,0,300,50,1,2000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A�I�i�k�N<bgsound src=wav/Phant008.wav loop=1>�ܥX�F�@�Ӥj�󳽡A���Ӫk�O100."

case 5
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�p�U��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�p�U��',"&info(9)&",'�ī~',0,0,100,302,302,0,100,30,1,500,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A�I�i�k�N<bgsound src=wav/Phant008.wav loop=1>�ܥX�F�@�Ӥp�U���A���Ӫk�O100."
case 6
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�Ѫ��' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('�Ѫ��',"&info(9)&",'�ī~',0,0,100,400,400,0,150,150,1,2000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A�I�i�k�N<bgsound src=wav/Phant008.wav loop=1>�ܥX�F�@���Ѫ�סA���Ӫk�O100."
case 7
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='����C' and �֦���=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,����,���s,��T��,����,sm,����,���O,��O,�ƶq,�Ȩ�,�|��) values ('����C',"&info(9)&",'�t��',0,0,100,1000,1000,0,10000,80000,1,200000,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���=" & info(9)
	end if
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A�I�i�k�N<bgsound src=wav/Phant008.wav loop=1>�ܥX�F�@�⭸��C�A���Ӫk�O100."

case else
banbianshu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�A<bgsound src=wav/Phant008.wav loop=1>�B��u�O�t�t�r�A���򳣨S�ܥX��."
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
sd(199)="<font color=CFE0B0>�i���ܳN�j"& banbianshu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>