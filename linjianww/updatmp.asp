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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim conn,rs,rsuser
zm=""
on error resume next
s=Request.QueryString("subject")
subject=request.form("subject")
comment=request.form("comment")
name=trim(request.form("name"))
jijing=trim(request.form("jijing"))
shihe=trim(request.form("shihe"))
if name<>"���w" then
set rs=server.createobject("adodb.recordset")
rs.open ("select * from �Τ� where �m�W='"&name&"'"),conn,0,1
if rs.eof then
name="���w"
Response.Redirect "../error.asp?id=441"
else
if rs("����")="�x��" then
zm=rs("����")
end if
end if
end if
if subject="" or (shihe<>"�k" and shihe<>"�k" and shihe<>"�Ҧ�") then
Response.Redirect "../error.asp?id=442"
else

if name<>"���w" then
conn.execute("update �Τ� set ����='"&subject&"',����='�x��',grade=5 where �m�W='"&name&"'")
end if
sql="Update ���� Set ����='" & subject & "',�A�X='" & shihe & " ',²��='" & comment & " ',�x��='" & name & " ',�������="&jijing&" Where ����='" & s & "'"
conn.execute(sql)
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','������s�ާ@')"
conn.Execute(sql)
rs.close
set rs=nothing
Response.Redirect "../ok.asp?id=701"
end if
%> 