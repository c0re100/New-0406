<%'�d�ݪ��A
function getstat(to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * FROM �Τ� WHERE �m�W='"&toname&"'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('�A�Ҭd�ݪ��H���s�b�I�I�I');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if toname=info(0) or info(2)>=9  then
	getstat=toname & "," & rs("�ʧO") & ",��Q:" & rs("�t��") & ",����:" &rs("����") & ",����:" &rs("����") &  ",�Z�\:" &rs("�Z�\") & ",���O:" & rs("���O") & ",��O:" & rs("��O") & ",�����O:" & rs("����") & ",���s�O:" & rs("���s")&",��eip:" & rs("ip") & ",�`�Uip:" & rs("regip") & ",�̫�ip:" & rs("lastip") & ",���A:" & rs("���A") & ",�H�c:" & rs("�H�c") & ",�y�O:" & rs("�y�O") & ",���l��:" & rs("���l��") & ",�D�w:" & rs("�D�w") & ",�޲z����:" & rs("grade") & ",�ֿn��:" & rs("allvalue") & ",��n��:" & rs("mvalue") & ",�n��:" & rs("�n��") & ",�Ȩ�:" & rs("�Ȩ�")
else
	if  info(2)>rs("grade") then
		Response.Write "<script language=JavaScript>{alert('���޲z�ŧO��A���A�Y�T�d�ݡI�I�I');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	getstat=toname & "," & rs("�ʧO") & ",��Q:" & rs("�t��") & ",����:" &rs("����") & ",����:" &rs("����")
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
