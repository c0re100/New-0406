<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=info(0)%>
<!--#include file="data1.asp"-->
<%sql="select * from �Ы� where ��D='" & my & "' or ��Q='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "�z�٨S���R�Фl�O�C"
location.href = "fangwu.asp"
</script>
<%else
if day(rs("�C�a"))<day(now()) and month(rs("�C�a"))<month(now()) and year(rs("�C�a"))<year(now()) and isnull(rs("�C�a")) then
sql="update �Ы� set �C�a=now() where ��D='"& my &"' or ��Q='" & my & "'"
set rs=conn.execute(sql)
set rs=nothing	
	conn.close
	set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from �Τ� where id="&info(9),conn
sql="update �Τ� set �y�O=�y�O+100,��O=��O+10000 where id="&info(9)
set rs=conn.execute(sql)
Response.Redirect "../ok.asp?id=602"
else
%><script language=vbscript>
MsgBox "�z���ѹC�L�a�F�ڡC"
location.href = "xiaowu7.asp"
</script>
<%end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>