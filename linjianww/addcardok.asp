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
wupinname=trim(Request.Form("wupinname"))
wupinlx=trim(Request.Form("wupinlx"))
wupinnl=abs(trim(Request.Form("wupinnl")))
wupintl=abs(trim(Request.Form("wupintl")))
wupingj=abs(trim(Request.Form("wupingj")))
wupinsm=trim(Request.Form("wupinsm"))
wupinfy=abs(trim(Request.Form("wupinfy")))
wupinyl=abs(Request.Form("wupinyl"))
id=trim(Request.Form("wupinID"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="insert into 物品買 (物品名,類型,內力,體力,攻擊,說明,防御,銀兩) values ('"&wupinname&"','"&wupinlx&"',"&wupinnl&","&wupintl&","&wupingj&",'"&wupinsm&"',"& wupinfy &","& wupinyl &")"
conn.Execute(sql)
sql="insert into 管理記錄 (姓名,時間,ip,記錄) values ('"&info(0)&"',now(),'"&info(15)&"','物品增加')"
conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%> 