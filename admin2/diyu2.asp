<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
name=info(0)
yilao=request.form("yilao")
if yilao<>"�L" then
'����Τ�
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,��O,��O�[,�Ȩ� from �Τ� where id="&info(9),conn
tlj=(rs("����")*1500+5260)+rs("��O�[")
if rs("��O")>tlj then
Response.Write "<script language=JavaScript>{alert('�A����O�w�g���F���ݭn�A�ɡI');location.href = 'diyu2.asp';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
If Rs.Bof OR Rs.Eof Then
mess="�A����i��v��"
else
sm=rs("��O")
select case yilao
case "�@�ŷҤ�"
bl=1.5
cd=10000-sm
case "�G�ŷҤ�"
bl=1
cd=10000-sm
case "�T���u��"
bl=0.5
cd=10000-sm
end select
fy=int(cd/bl)
if fy>rs("�Ȩ�") then
mess="�A�w�g�a���s�ϩR�����S���F�C�ٷQ�Ҥ��@�H�I"
else
conn.execute "update �Τ� set ��O=10000,�Ȩ�=�Ȩ�-" & fy & " where id="&info(9)
mess="�g�L�a�����L�������ڿN�A�A��F" & fy & "��Ȥl�ͩR��_��10000�I"
end if
end if
else
mess="�A�ݭn�Һ��F"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('"& mess &"');location.href = 'diyu.asp';}</script>"
response.end%>
