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
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../wb/ljjmwb.asa")
Conn.Open connstr
sql="select qdlm from bar "
Set Rs=conn.Execute(sql)
if rs("qdlm")=true then
conn.Execute "update bar set qdlm=false where id="&request("wbid") 
else
conn.Execute "update bar set qdlm=true where id="&request("wbid")
end if 
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','加盟網吧確認操作')"
conn.Execute(sql)
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Redirect "../ok.asp?id=700"
%> 