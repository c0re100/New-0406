<%
'�����M
function jiuqingdao(fn1,to1,toname)
if toname="�j�a" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('�ާ@�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<20 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n20�ťH�W�~�i�H�ϥε����M�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O,�ާ@�ɶ� FROM �Τ� WHERE id="&info(9),conn
fla=rs("�k�O")
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');}</script>"
	Response.End
end if
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o1000�I�ڡI');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-1000,�ާ@�ɶ�=now() where id="&info(9)
rs.close
rs.open "select ��O,���� FROM �Τ� WHERE �m�W='"&toname&"'",conn
tl=rs("��O")
money=int(rs("��O")/3)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�����M' and �ƶq>=1 and �֦���=" & info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A�������M�ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update ���~ set �ƶq=�ƶq-1 where id="& id &" and �֦���=" & info(9)
r=int(rnd*2)+1
select case r
case 2
conn.execute "update �Τ� set ��O=��O-"&money&" where �m�W='"&toname&"'"
jiuqingdao=info(0) & "�q�y�����X�@�⵴���M<bgsound src=wav/fx42.wav loop=1>�A<img src='img/9.gif' width='44' height='44'>�c�����a���<font color=red>"&towho&"</font>��F�U�h,�����a�A"&towho&"�ڪ��@�n�A�Q�뤤�߯�,���g�ɮ��������ڡA��O���h1/3�A�@�p"&money&"�I......" 
if tl<=0 then 
conn.execute "update �Τ� set ���A='��' where �m�W='"&toname&"'"
jiuqingdao=info(0) & "�q�y�����X�@�⵴���M<bgsound src=wav/ZR2199.wav loop=1>�A<img src='img/9.gif' width='44' height='44'>�c�����a���<font color=red>"&towho&"</font>��F�U�h,�����a,"&towho&"�ڪ��@�n�A�ѩ���O����A�Q����릺�A�ۥj�h���žl���..." 
 call boot(towho)
end if
case else
jiuqingdao=info(0) & "�q�y�����X�@�⵴���M<bgsound src=wav/fx42.wav loop=1>�A<img src='img/9.gif' width='44' height='44'>�c�����a���<font color=red>"&towho&"</font>��F�U�h,�����a,��ƬO�R��`�B�߫�֡A�S����......"
end select
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>