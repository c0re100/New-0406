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
id=request("id")
leixin=trim(request.form("leixin"))
jiedao=trim(request.form("jiedao"))
money=trim(request.form("money"))
sn=trim(request.form("sn"))
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../CHANGAN/fangzi.mdb")
Conn.Open connstr

if Request.Form("submit")="�R��" then
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))
	conn.Execute "update �Ы� set ��D='�L',��Q='�L' where id="&nameid
end if
if Request.Form("submit")="�K�[" then
sql="select * from �Ы� "
set rs=Conn.execute(sql)
for i=1 to sn
conn.Execute "insert into �Ы�(����,���,��m,��D,��D,��Q) values ('" & leixin & "',"&money&",'���w','" & jiedao & "','�L','�L')"
next
	end if
set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�ާ@���\�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
