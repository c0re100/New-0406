<%
function dmjwp(to1)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="duju/dmj.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'�p�G�A���A�Ⱦ����jet4.0�A�ШϥΤU�����s����k�A�����{�ǩʯ�
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 

sql="select oldmj from dmj where uto='"& info(0) & "' and ufrom='"& to1 & "'"
Set Rs=conn.Execute(sql)
if rs.eof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('�A�S����[ " & to1 & " ]���±N!');location.href = 'javascript:history.go(-1)';</script>"
  response.end
end if
oldmj=rs("oldmj")
	rs.close
	dmjwp=info(0) & ":" &to1&"��~���F�@�i " & showMj(oldmj) &"!"
	conn.close
	set rs=nothing
	set conn=nothing
	end Function

%>
<script language="JavaScript">parent.f3.location.reload();</script>
<%
function showMj(Mj)
showMj="<input type=image border=0 src=""duju/mj/" & intMjp(Mj) & ".gif"" onclick=""javascript:parent.f3.location.href='duju/dmj-xp.asp';"" >"
end function
%>