<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
username=info(0)
grade=info(3)
if grade<5 then
	Response.Write "<script language=JavaScript>{alert('����p���A���ɬ��|���n�S�W�檺�h�Q!');location.href = 'index.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT �ʧO,�y�O FROM �Τ� WHERE �m�W='"&username&"'"
set Rs=conn.Execute(sql)
sex=rs("�ʧO")
meimao=rs("�y�O")
if sex<>"�k" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���S���d���r�A�ɬ��|�̥i���n�k���@!');location.href = 'index.asp';}</script>"
	response.end
end if
if meimao<1000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���S���d���r�A�A�����o���٨ӳo��!');location.href = 'index.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from �W�� where �m�W='" & username & "'"
set Rs=conn.Execute(sql)
if rs.bof or rs.eof then
sql="insert into �W��(�m�W,������,����) values ('" & username & "'," & meimao & ",'�L')"
conn.execute(sql)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT id FROM �Τ� WHERE �m�W='"&username&"'"
set Rs=conn.Execute(sql)
sql="update �Τ� set �Ȩ�=�Ȩ�+1000000 where �m�W='"& username &"'"
conn.Execute sql
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ߡA�A�����������|���h�Q�A�o��樭�Ȩ�1000000�U�I!');location.href = 'index.asp';}</script>"
	response.end
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�w�g�O���ɬ��|���h�Q�F�A����٨ӵn�O�r�I!');location.href = 'index.asp';}</script>"
	response.end
connt.close
end if
%> 