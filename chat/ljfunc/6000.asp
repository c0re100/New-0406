<%
'�D���ݤR
function kk(fn1,to1,toname)
if toname="�j�a" or toname=info(0) or toname=Application("ljjh_automanname")  then
	Response.Write "<script language=JavaScript>{alert('�ﹳ�����A�ЬݥJ�ӤF�I�I');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('���ŤӧC�A�n10�ťH�W��i�ΡI');}</script>"
	Response.End
end if
conn.execute "update �Τ� set ���O=���O+100000 where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
kk=�i�D���ݤR�jinfo(0) & "�ۤ߬�D���@��i���dñ�j,�� ��R���� �����һ��D:�o�O�@�䰷�dñ.��O�U�U��.�Ҥ�W���A���~�����λݭn�n�n��i.���i�L���޳�.�n���M���O�|�C�C������"&info(0) & �Ǳ§A2000�I���O���A����w�d.....
"" 
else
conn.execute "update �Τ� set ��O=��O-100000 where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
kk=�i�D���ݤR�jinfo(0) & "�ۤ߬�D���@��i���Bñ�j�G�� ��R���� �����һ��D:�o�O�@�侵�Bñ.�D�`�D�`�����n.�ݨӧA�O���B���k�F����R�v�u����ӧA����O�Ӭ��A�Ѱ����B��"&info(0) & �����A��O2000.....

"" 
else
conn.execute "update �Τ� set ����=����+200000 where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
kk=�i�D���ݤR�jinfo(0) & "�ۤ߬�D���@��i�o�]ñ�j�G�� ��R���� �����һ��D:�o�O�@��o�]ñ�۷��n.���Ҥ�W�һ��A���ް�����]�B���|��o��.�u�O���ߧA�F�o�O�W��ñ��"&info(0) & �����K�ذe�A20�U��.....
"" 
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>