<%'�ɫ¤O��baodou
function baodou()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select ����,grade,�w���I��,�ɨ��ɶ� FROM �Τ� WHERE id="&info(9),conn
if trim(rs("����"))="�@�k" and rs("grade")=3 then
	doudi=500
else
	doudi=1000
end if
if rs("�w���I��")<doudi then
	Response.Write "<script language=JavaScript>{alert('�A���̦��¤O���H1�Ө��i�H�b�@�p���¤O��3���I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sj=DateDiff("n",rs("�ɨ��ɶ�"),now())
if sj<=60 then
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('"& info(0) & "�¤O�ɶ����L�A�ٳ�"& ss &"����!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update �Τ� set �ɨ��ɶ�=now(),�w���I��=�w���I��-"& doudi &" where id="&info(9)
baodou=info(0) & "���P�A�ɫ¤O���ާ@���\�A�q�{�b�}�l�A�������O�|���Ӥj3���I"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>