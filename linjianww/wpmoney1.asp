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
bei=abs(request.form("moneybei"))
lx=trim(request.form("lx"))
if bei>=0 then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	conn.Execute "update 物品買 set 銀兩=(abs(內力)+abs(體力)+abs(攻擊)+abs(防御))*"& bei &"  where 類型='"& lx &"'"
	conn.close
	set conn=nothing	
Response.Redirect "../ok.asp?id=700"
end if
%> 