<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
sql="SELECT ���� FROM �Τ� WHERE �m�W='" & name & "'"
Set Rs=conn.Execute(sql)
jishu=rs("����")
sql="select ������ from ���� where �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "�A�S�����ڡA�O���O�Q�F�ڡA�֥h���]�@���a�C"
location.href = "qmg.htm"
</script><%
conn.close
response.end
else
mw=rs("������")
tl=mw*8
conn.execute("update ���� set �ɭ�=true where �֦���='"&name&"' and �ɭ�=false")
conn.execute("update �Τ� set ��O=��O+"&tl&" where �m�W='"&name&"'")
Response.write"�A��A�ˤⰵ������F�A��O�W�[" & tl & "�I�C"
set rs=nothing
conn.close
set conn=nothing
end if	
%> 