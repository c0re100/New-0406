<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
my=info(0)
sql="select �Ȩ� from �Τ� where �m�W='" & my & "'"
set rs=conn.execute(sql)
xianjin=rs("�Ȩ�")
sql="SELECT ���,��O,�֦��� FROM �歱 where ID=" & id & " and �ɭ�=false"
Set rs=conn.Execute(sql)
if not(rs.eof and rs.bof) then
yin=rs("���")
tl=rs("��O")
yy=rs("�֦���")
if yin <= xianjin then
sql="update �Τ� set �Ȩ�=�Ȩ�-" & yin & ",��O=��O+" & tl & "   where  �m�W='" & my & "'"
conn.execute(sql)
conn.execute("update �歱 set �ɭ�=true where ID="&id)
response.write "�A��F�����]�v��"& yy &"�������A�W�[��O"&tl
else
response.write "����𭱡A��]�G�A���Ȩ⤣���F"
conn.close
response.end
end if
end if
set rs=nothing
conn.close
set conn=nothing
%> 