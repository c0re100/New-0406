<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from ��� where ����='����' and �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "�A�S�����ơA��򰵭��ڡA�֨�楫���h�R�a�C"
location.href = "caichang.asp"
</script><%conn.close
response.end
else
randomize timer
mw=2*rnd*rs("������")
sql="select * from ���� where  �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="select * from ���� where  �ɭ�=true"
if rs.eof or rs.bof then		
sql="insert into ���� (�֦���,������) values ('"& name &"'," & mw & ")"
rs=conn.execute(sql)
else
sql="update ���� set �ɭ�=false,������="&mw&",�֦���='" & name & "' where �ɭ�=true"
rs=conn.execute(sql)
end if
sql="update ��� set �ɭ�=true where ����='����' and �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
%><script language=vbscript>
MsgBox "�A��������i��̡A�@�J�n���N�֦b���e�F�C"
location.href = "m2.htm"
</script><%else
sql="update ���� set �ɭ�=true where �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
%>
<script language=vbscript>MsgBox "�A��H�e�S�����n�����˱��A�^��p�Э��s�ӹL�C"
location.href = "ctl_work_noodle.htm"
</script><%
set rs=nothing
conn.close
set conn=nothing
end if	
end if
%> 