<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
num=abs(trim(request.form("text")))
if not isnumeric(num) then Response.Redirect "../error.asp?id=54"
if num>1000000 then
 	Response.Write "<script language=JavaScript>{alert('�@���|�O�Ȥ���Ӥj!');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �|������,�|���O from �Τ� where id="&info(9),conn
if rs("�|������")=0 then
Response.Write "<script language=JavaScript>{alert('�A���O�|���Ц^�a�I');location.href = '../help/info.asp?type=2&name=�|���[�J��k';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if rs("�|���O")<num then
Response.Write "<script language=JavaScript>{alert('�A������h�|�O�ܡH');location.href = 'hxiehy.asp';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
jbb=rs("�|���O")-num
jb=int(num*100)
conn.execute("update �Τ� set �|���O="&jbb&",����=����+"&jb&",�ާ@�ɶ�=now()  where id="&info(9))
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.write "<script language='javascript'>alert ('�|�O�ഫ���������\�A�Ъ`�N�d��!');location.href='hxiehy.asp';</script>"
Response.End
%>
