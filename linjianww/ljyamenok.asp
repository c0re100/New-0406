<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
a=request("a")
shenfen=trim(request.form("shenfen"))
shenfen1=trim(request.form("shenfen1"))
name1=trim(request.form("name1"))
id=request("id")
dj=trim(request.form("dj"))
dj1=trim(request.form("dj1"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if a="m" then
rs.open "SELECT ���� FROM �Τ� where id="&id,conn
if rs.bof or rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.write "�藍�_�A�S���ӥΤ�I"
	response.end
end if
if rs("����")="������" then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('���H�O�������A����ק�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if shenfen="������" or dj=11 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('����������@�Ӥ��o�ק�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if info(2)=10 and dj>=10 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))
	conn.Execute "update �Τ� set ����='" & shenfen & "', grade=" & dj & " where id="&nameid
end if
if Request.Form("submit")="�}��" then
	nameid=int(abs(request("id")))

	conn.Execute "update �Τ� set ����='�C�L',����='�̤l', grade=1 where id="&nameid
end if
if Request.Form("submit")="�K�[" then
if shenfen1="������" or dj1=11 then
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�������V�å��������঳�@�ӡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
if info(2)=10 and dj1>=10 then
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.open "SELECT id FROM �Τ� where �m�W='"&name1&"'",conn
nameid1=rs("id")
	conn.Execute "update �Τ� set ����='�x��',����='" & shenfen1 & "', grade=" & dj1 & " where id="&nameid1
end if
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�x���H���ק�ާ@')"
conn.Execute(sql)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
