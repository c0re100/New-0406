<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=request("id")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT ���~�W,�Ȩ�,����,����,�\��,������ FROM ��� where ID=" & id
Set Rs=conn.Execute(sql)
wu=rs("���~�W")
yin=rs("�Ȩ�")
lx=rs("����")
sm=rs("����")
gn=rs("�\��")
mw=rs("������")
sql="select �Ȩ� from �Τ� where �m�W='" & my & "'"
rs=conn.execute(sql)
if yin<=rs("�Ȩ�") then
sql="update �Τ� set �Ȩ�=�Ȩ�-" & yin & "  where �m�W='" & my & "'"
rs=conn.execute(sql)
sql="select * from ��� where ���~�W='" & wu & "' and �ɭ�=false and �֦���='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="select * from ��� where ���~�W='" & wu & "' and �ɭ�=true"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
sql="insert into ���(���~�W,�֦���,����,����,�\��,������,�Ȩ�) values ('"& wu &"','"& my &"','"& lx &"','" & sm &"','"& gn &"',"& mw &","& yin &")"
rs=conn.execute(sql)
conn.close
Response.Redirect "caichang.asp"
else
id2=rs("id")
sql="update ��� set �֦���='"&my&"',�ɭ�=false where id=" & id2
rs=conn.execute(sql)
conn.close
Response.Redirect "caichang.asp"
end if
else
response.write "�ѩ�A�w�ʶR�F�o�Ӫ��~�A�ҥH����A�R�I"
conn.close
response.end
				
end if
else
response.write "����R�F��A��]�G�A���Ȩ⤣���F"
conn.close
response.end
end if
rs.close
set rs=nothing
set conn=nothing
%> 