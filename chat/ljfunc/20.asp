<%'�U��
function cefen(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�U�ʹﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(2)<2 then
	Response.Write "<script language=JavaScript>{alert('�޲z�ŤӧC�A�n2�ťH�W�~�i�H�U�ʹ��I');}</script>"
	Response.End
end if
fn1=trim(fn1)
if len(fn1)>10 then
	Response.Write "<script language=JavaScript>{alert('�U�ʨ������פ��i�j��10�Ӧr�šI');}</script>"
	Response.End
end if
if instr(fn1,"�x��")<>0 then
	Response.Write "<script language=JavaScript>{alert('�L�@�x���A�@����H');}</script>"
	Response.End
end if
if len(fn1)>2 and (instr(fn1,"����")<>0 or instr(fn1,"�@�k")<>0 or instr(fn1,"��D")<>0 or instr(fn1,"�Ƥ�")<>0 ) then
	Response.Write "<script language=JavaScript>{alert('����:���ѡB�@�k�B��D���t�ΫO�d�A�Ф��n�ϥΦb�W�����I');}</script>"
	Response.End
end if
cefeng1=instr(says,"=")
cefeng2=instr(says,",")
cefeng3=instr(says,"grade")
cefeng4=instr(says,"����")
cefeng5=instr(says,"����")
cefeng6=instr(says,"�x��")
cefeng6=instr(says,"�����`��")
cefeng6=instr(says,"����")
cefeng6=instr(says,"�Ư���")
cefeng7=instr(says,"`")
cefeng8=instr(says,"\")
cefeng9=instr(says,"<")
cefeng10=instr(says,">")
if cefeng1<>0 or cefeng2<>0 or cefeng3<>0 or cefeng4<>0 or cefeng5<>0 or cefeng6<>0 or cefeng7<>0 or cefeng8<>0  or cefeng9<>0 or cefeng10<>0 then
	Response.Write "<script language=JavaScript>{alert('�ާ@���~�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ���� from �Τ� where id="& info(9),conn
aaa=rs("����")
rs.close
rs.open "select �m�W,grade,����,�|������ from �Τ� where �m�W='"&toname&"' and ����='" & aaa& "'",conn
toname=rs("�m�W")
grade=rs("grade")
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('["& toname &"]�]���O�A�����̤l�A�Q�@����H');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if info(2)<grade then
	Response.Write "<script language=JavaScript>{alert('["& toname &"]��������A���ڡI');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
select case fn1
case "�Ƥ�"
if info(2)<5 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���޲z�Ť����ާ@�I');}</script>"
	Response.End
end if
	if rs("����")<25 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]�٤���25�šA����ʬ��ƴx��!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where grade=3 and ����='�ƴx��' and ����='"&info(5)&"'")
	musers=tmprs("�ƶq")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=2 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]�{�b�A�����ƴx����2�ӤF�A���n�A�ʤF�I');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �Τ� set ����='�ƴx��',grade=3 where �m�W='"&toname&"'"
fn1="�ƴx��"
case "����"
if info(2)<4 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���޲z�Ť����ާ@�I');}</script>"
	Response.End
end if
	if rs("����")<15 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]�٤���15�šA����ʪ���!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if

	tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where grade=2 and ����='����' and ����='"&info(5)&"'")
	musers=tmprs("�ƶq")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=4 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]�{�b�A�������Ѧ�4�ӤF�A���n�A�ʤF�I');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �Τ� set ����='����',grade=2 where �m�W='"&toname&"'"

case "�@�k"
if info(2)<3 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���޲z�Ť����ާ@�I');}</script>"
	Response.End
end if
	if rs("����")<10 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]�٤���10�šA������@�k!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if

	tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where grade=2 and ����='�@�k' and ����='"&info(5)&"'")
	musers=tmprs("�ƶq")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=8 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]�{�b�A�����@�k��8�ӤF�A���n�A�ʤF�I');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �Τ� set ����='�@�k',grade=2 where �m�W='"&toname&"'"

case "��D"
if info(2)<2 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���޲z�Ť����ާ@�I');}</script>"
	Response.End
end if
	if rs("����")<5 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]�٤���5�šA����ʰ�D!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
tmprs=conn.execute("Select count(*) As �ƶq from �Τ� where grade=2 and ����='��D' and ����='"&info(5)&"'")
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
	if musers>=12 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]�{�b�A������D��12�ӤF�A���n�A�ʤF�I');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update �Τ� set ����='��D',grade=2 where �m�W='"&toname&"'"

case else
if aaa="�x��" and info(2)>=10 then
conn.execute "update �Τ� set ����='"& fn1 &"' where �m�W='"&toname&"'"
else
	conn.execute "update �Τ� set ����='"& fn1 &"',grade=1 where �m�W='"&toname&"'"
end if
end select
cefen=aaa&"���x���G"& info(0) & "�U��" & toname & "��" & aaa & "��<font color=red><b>" & fn1 &"</b></font>���\,�j�a���P�I"
'�O��
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& info(0) &"','"& toname &"','�U��','"& cefen & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>