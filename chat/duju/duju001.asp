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
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(5)<>0 then
Response.Write "<script Language=Javascript>alert('�n��սХ��h�j�U�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
'f=Minute(time())
'if f>30 then
'	Response.Write "<script language=JavaScript>{alert('�{�b�����աI��ծɶ����C�p�ɪ��e30����,�Ҧp�G17:00�}�l17:30����!!');}</script>"
'	Response.End 
'end if
if Request.Form("�Ȩ⧽")="�Ȩ⧽�@��" then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select �s�� FROM �Τ� WHERE id=" & info(9),conn
	if rs("�s��")<20000000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A���s�ڤ���2000�U�A����@���I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 �m�W FROM �b�u��� WHERE ����='���a'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('�{�b["&rs("�m�W")&"]�����a,�A����@���I�I');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into �b�u���(�m�W,����,�Ȩ�,��j�p,�ɶ�) values ('"&info(0) &"','���a',0,'��',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application.Lock
	Application("ljjh_dubozhang")=info(0)
Application.UnLock
	dubo="<font color=green>�i�Ȩ�䧽�@���j</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>�b�F�C����b�u����ӽа������\�A�{�b�j�a�i�H����@��F,�֨ӧ֨ӧr  "
 
end if

if Request.Form("�k�O��")="�k�O���@��" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select �k�O FROM �Τ� WHERE id=" & info(9),conn
	if rs("�k�O")<5000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A���k�O����5000�A����@���I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 �m�W FROM �k�O�䧽 WHERE ����='���a'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('�{�b["&rs("�m�W")&"]�����a,�A����@���I�I');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into �k�O�䧽(�m�W,����,�k�O,��j�p,�ɶ�) values ('"&info(0) &"','���a',0,'��',now())"	
	rs.close
Application.Lock
	Application("ljjh_dubozhang2")=info(0)
Application.UnLock
	dubo="<font color=green>�i�k�O�䧽�@���j</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>�b�F�C����k�O�䧽�ӽа������\�A�{�b�j�a�i�H�U�̹B���@��F,�֨ӧ֨ӧr  "
end if

if Request.Form("�w�I��")="�w�I���@��" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select allvalue FROM �Τ� WHERE id=" & info(9),conn
	if rs("allvalue")<3000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A���I�Ƥ���3000�I�A���৤���I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 �m�W FROM �I�ƽ䧽 WHERE ����='���a'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('�{�b["&rs("�m�W")&"]�����a,�A���৤���I�I');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into �I�ƽ䧽(�m�W,����,�I��,��j�p,�ɶ�) values ('"&info(0) &"','���a',0,'��',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application.Lock
	Application("ljjh_dubozhang3")=info(0)
Application.UnLock
	dubo="<font color=green>�i�w�I�䧽�@���j</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>�b�F�C����s�I�ƽ䧽�ӽа������\�A�{�b�j�a�i�H�U�̹B���@��F,�֨ӧ֨ӧr  "
end if

if Request.Form("������")="�������@��" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select ���� FROM �Τ� WHERE id=" & info(9),conn
	if rs("����")<4000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A����������4000�ӡA���৤���I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 �m�W FROM �����䧽 WHERE ����='���a'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('�{�b["&rs("�m�W")&"]�����a,�A���৤���I�I');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into �����䧽(�m�W,����,����,��j�p,�ɶ�) values ('"&info(0) &"','���a',0,'��',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application.Lock
	Application("ljjh_dubozhang4")=info(0)
Application.UnLock
	dubo="<font color=green>�i�����䧽�@���j</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>�b�F�C��������䧽�ӽа������\�A�{�b�j�a�i�H�U�̹B���@��F,�֨ӧ֨ӧr  "
end if

if Request.Form("�[��")="�U�䧽�[��" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
rs.open "select �ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
tmprs=conn.execute("Select count(*) As �ƶq from �b�u��� where �Ȩ�<>0")
	durs1=tmprs("�ƶq")
set tmprs=nothing	
	rs.open "select top 1 �m�W FROM �b�u��� WHERE ����='���a' and �m�W='"&info(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A�]���O���A�A�Q�d����H�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
tmprs=conn.execute("Select count(*) As �ƶq from �k�O�䧽 where �k�O<>0")
	durs2=tmprs("�ƶq")
set tmprs=nothing
	rs.open "select top 1 * FROM �k�O�䧽 WHERE ����='���a' and �m�W='"&info(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A�]���O���A�A�Q�d����H�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close

	tmprs=conn.execute("Select count(*) As �ƶq from �I�ƽ䧽 where �I��<>0")
	durs3=tmprs("�ƶq")
set tmprs=nothing
	rs.open "select top 1 �m�W FROM �I�ƽ䧽 WHERE ����='���a' and �m�W='"&info(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A�]���O���A�A�Q�d����H�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
conn.Execute "update �Τ� set �ާ@�ɶ�=now() where id="&info(9)
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<font color=green>�i�䧽�[�ܡj</font><img src='../jhimg/sz.gif'>���a:<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�s��G�R�o�hĹ�o�h�A�R�o��Ĺ�o�֡I�R�F�z�ᮬ(��򤣦h�R�I�r!)���R�z��ᮬ(�n�R�N���F)�A<font color=red>�i�Ȩ�䧽�j</font>�٦�:"& (10-durs1) &"�ӤU�`��}��  <font color=red>�i�k�O�䧽�j</font>�٦�:"& (5-durs2) &"�ӤU�`��}��  <font color=red>�i�w�I�䧽�j</font>�٦�:"& (5-durs3) &"�ӤU�`��}��  "
end if
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
sd(194)=info(0)
sd(195)=info(0)
sd(199)="<font color=#FFC01F>"& dubo &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{location.href = 'javascript:history.go(-1)';}</script>"
Response.End 
%>