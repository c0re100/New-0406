<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
	sex=request.form("sex")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 銀兩,領錢,體力 FROM 用戶 where id="&info(9)&" and 性別= '" &sex&"'",conn
if rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對不起，你不是江湖中人');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if
if rs("銀兩")<380 then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('對不起，你的錢不夠');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if
if day(rs("領錢"))<day(now()) or month(rs("領錢"))<month(now()) or year(rs("領錢"))<year(now()) or isnull(rs("領錢")) then
energy=rs("體力")
tempdate=date
energytemp=energy+1000
conn.execute "update 用戶 set 銀兩=銀兩-300,領錢='"&tempdate&"',體力='"&energytemp&"' where id="&info(9)
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('你洗完了溫泉浴感覺輕松多了。');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
else
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('對不起，你已經洗過了');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if
%>
