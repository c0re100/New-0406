<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
sql="select ������ from ��� where ����='����' and �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>MsgBox "�A�S��9�z�A��򰵭��ڡA�֨�楫���h�R�a�C"
location.href = "caichang.asp"
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
sql="update ��� set �ɭ�=true where ����='����' and �֦���='" & name & "' and �ɭ�=false"
set rs=conn.execute(sql)
%>
<script language=vbscript>MsgBox "�A��9�z��i��̡A�@�J�n���N�֦b���e�F�C"
location.href = "m3.htm"
</script><%
set rs=nothing
conn.close
set conn=nothing
end if	
end if	
%> 