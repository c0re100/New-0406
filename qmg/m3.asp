<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select ������ from ��� where ����='�D��' and �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "�A�S���z���A�֥h����R�a�C"
location.href = "ciachang.asp"
</script><%conn.close
response.end
else
randomize timer
mw=2*rnd*rs("������")
sql="select * from ���� where  �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "�A�٨S�����Ĥ@�B�ڡC"
location.href = "m1.htm"
</script><%else
sql="update ���� set ������=������+" &mw& " where �֦���='" & name & "' and �ɭ�=false"
rs=conn.execute(sql)
sql="update ��� set �ɭ�=true where ����='�D��' and �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
%>
<script language=vbscript>MsgBox "�A�� ����i��̡A�@�J�n���N�֦b���e�F�C"
location.href = "m4.htm"
</script><%
set rs=nothing
conn.close
set conn=nothing
end if	
end if	
%> 