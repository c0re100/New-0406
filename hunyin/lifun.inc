<%
Response.Buffer=true
name1=Request.form("name12")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name22")
name2=trim(name2)
name2=server.HTMLEncode(name2)
pass=request.form("pass2")
pass=trim(pass)
mess=Request.form("mess2")
mess=trim(mess)
mess=server.HTMLEncode(mess)
if name1="" or name2="" then
message="�m�W���ର��"
elseif pass="" then
message="�K�X���ର��"
elseif mess<0 or mess>1 then
message="�]�����μƥؤ���"
else
%><!--#include file="../data.asp"--><%
'����Τ� �y�O�j��100�A���j��1000
sql="SELECT * FROM �Τ� WHERE �m�W='" & name1 & "' and �K�X='" & pass & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
response.write "����n�O�I���ˬd�A���m�W�M�K�X"
conn.close
response.end
else
xb1=trim(rs("�ʧO"))
if rs("�Ȩ�")>=1000 then
rs.close
set rs=nothing
set rs=conn.execute("SELECT * FROM ��� WHERE trim(�m�W1)='" & name1 & "'trim(�m�W2)='" & name2 & "'")
if not (rs.bof or rs.eof) then

sql="update �Τ� set �Ȩ�=�Ȩ�-5000 where �m�W='" & name1& "'"
conn.execute sql
sz = "'" & name1 & "','" & name1& "','" & mess & "', now()"
into_db = "INSERT INTO ��� (�m�W1, �m�W2, ����, �ɶ�) VALUES(" & sz & ")"
conn.Execute(into_db)
sql="delete * from ��� where now()-�ɶ�>10"
conn.execute sql
message="�n�O���\�I�����O��5000��I"
else %>
<script language=vbscript>
MsgBox "�A�S���ѱC�ڡH�d���F�a�I"
location.href = "../jh.asp"
</script>

<%               end if
else %>
<script language=vbscript>
MsgBox "�A���̦�5000��ڡA�S���ٷQ���B�H���ڥh�a�I"
location.href = "../jh.asp"
</script>

<%		end if
end if
conn.close
end if
%>
<script language=vbscript>
MsgBox "<%=message%>"
location.href = "../jh.asp"
</script>
 