<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")

username=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT �ʧO,�Ȩ� FROM �Τ� WHERE �m�W='"&username&"'"
set rs=conn.execute(sql)
sex=rs("�ʧO")
yl=rs("�Ȩ�")
if sex="�k" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���S���d���r�A�ɬ��|�̥i���n�k���@!');location.href = 'index.asp';}</script>"
	response.end
end if
if yl<5000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A���S���d���r�A�S���]�Qū��!');location.href = 'index.asp';}</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from �W�� where �m�W='" & username & "'"
set rs=conn.execute(sql)
if not rs.eof then
sql="delete from �W�� where �m�W='"& username &"'"
conn.execute sql
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT id FROM �Τ� WHERE �m�W='"&username&"'"
set rs=conn.execute(sql)
sql="update �Τ� set �Ȩ�=�Ȩ�-5000000 where �m�W='"& username &"'"
set rs=conn.execute(sql)
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ߡA�A�ש����}�F�C�ӡA�ٲM�U�ڻȨ�500�U!');location.href = 'index.asp';}</script>"
	response.end
else
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�w�g���O���ɬ��|���h�Q�F�A����٨�!');location.href = 'index.asp';}</script>"
	response.end
end if
%> 