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

if Request.Form("submit")="刪除" then
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('你沒有這個權限！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
	nameid=int(abs(request("id")))
	conn.Execute "update 房屋 set 戶主='無',伴侶='無' where id="&nameid
end if
if Request.Form("submit")="添加" then
sql="select * from 房屋 "
set rs=Conn.execute(sql)
for i=1 to sn
conn.Execute "insert into 房屋(類型,售價,位置,街道,戶主,伴侶) values ('" & leixin & "',"&money&",'長安','" & jiedao & "','無','無')"
next
	end if
set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('操作成功！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
%>
