<%
'�t���_��
function peibashi(fn1,to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
rs.open "select ���� FROM �Τ� WHERE id="&info(9),conn
if rs("����")<10 then
Response.Write "<script language=JavaScript>{alert('���\��ݭn10�ž԰����ŧr�I');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM ���~  WHERE ���~�W='���_��'   and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A�����B��B���_�۶ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM ���~  WHERE  ���~�W='���_��'  and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A�����B��B���_�۶ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select id FROM ���~  WHERE  ���~�W='���_��'  and �ƶq>0 and �֦���="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('�A�����B��B���_�۶ܡH');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update ���~ set �ƶq=�ƶq-1 where ���~�W='���_��'  and �֦���="&info(9)
conn.execute "update ���~ set �ƶq=�ƶq-1 where ���~�W='���_��'  and �֦���="&info(9)
conn.execute "update ���~ set �ƶq=�ƶq-1 where ���~�W='���_��'  and �֦���="&info(9)
rs.close
rs.open "select id FROM ���~  WHERE ���~�W='�]�O�p��'  and �ƶq>0 and �֦���="&info(9)
If Rs.Bof OR Rs.Eof then
		conn.execute "insert into ���~(���~�W,�֦���,����,���O,��O,�ƶq,�Ȩ�,����,sm,����,����,���s,�|��) values ('�]�O�p��',"&info(9)&",'�k��',0,0,1,200000,1100,1100,0,0,0,"&aaao&")"
	else
id=rs("id")
		conn.execute "update ���~ set �ƶq=�ƶq+1,�|��="&aaao&" where id="& id &" and �֦���="&info(9)
	end if
peibashi=info(0) & "���X���B��B�ŤT���_�ۡA�T���_�۵��X�b�@�_�A�@�D���~�ɰ_�A<img src='img/look52.gif'>�T���_�ۤƦ��F�]�O�p��." 

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>