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
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
dim conn,rs,rsuser
on error resume next
s=Request.QueryString("username")
rs=conn.execute("select * from ���� where ����='"&s&"'")
zm=rs("�x��")
rs.close
set rs=nothing
set rs=server.createobject("adodb.recordset")
rs.open ("select * from �Τ� where �m�W='"&zm&"'"),conn,0,1
if rs.bof or rs.eof then
Response.Redirect "../error.asp?id=446"
else
conn.execute("update �Τ� set ����='�̤l',grade=1 where �m�W='"&zm&"'")
end if

conn.execute("update ���� set �x��='���w' where ����='"&s&"'")
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�x���}��')"
conn.Execute(sql)
rs.close
set rs=nothing
Response.Redirect "../ok.asp?id=703"
%> 