<%
'�ͤ�J�|
function shendangao(fn1,to1)
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W�~�i�H�ާ@�I');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �k�O FROM �Τ� WHERE id="&info(9),conn
fla=rs("�k�O")
if fla<100 then
Response.Write "<script language=JavaScript>{alert('�A���k�O�����L�k�I�i�r�A�ܤ֤]�o100�I�ڡI');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update �Τ� set �k�O=�k�O-100 where id="&info(9)
rs.close
rs.open "select id FROM ���~ WHERE ���~�W='�ͤ�J�|' and �ƶq>2 and �֦���=" & info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A��2�ӥͤ�J�|�ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update ���~ set �Ȩ�=200000,�ƶq=�ƶq-2 where ���~�W='�ͤ�J�|' and �֦���=" & info(9)
conn.execute "update �Τ� set ��O=��O+1000000 where id="&info(9)
shendangao=info(0) & "���a�q�a�W�߰_�@�⳷���g�M��ۤv���{�l��F�}�Ӷ�F�@���k���ͤ�J�|<img src='pic/dz59.gif'>�i�h�A�K�K�A�ӧa�A���ڧr�A�����ƥ����ڰڡA<font color=red>["&info(0)&"]</font>����O�ɺ�100�U�I." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>