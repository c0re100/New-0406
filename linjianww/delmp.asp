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
dim conn,rs,rsuser
owner=request.querystring("owner")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
select case request.querystring("action")
case "delete"
set rstmp=conn.execute("Select * From ���� Where ����='" & owner & "' ")
if rstmp.eof then
Response.Redirect "../error.asp?id=447"
else
conn.execute "Delete  From ���� Where ����= '" & owner & "' "
cname=rstmp("�x��")
conn.execute("Update �Τ� set ����='�C�L',����='�̤l',grade=1 where ����='"&owner&"'")
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�����R��')"
conn.Execute(sql)
Response.Redirect "../ok.asp?id=704"
end if
end select
set rs=nothing
%>
<script language=vbscript>
MsgBox "<%=response.write(str)%>"
location.href = "adminmpp.asp"
</script>