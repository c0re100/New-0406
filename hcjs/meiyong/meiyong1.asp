
<body bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩,魅力,狀態 from 用戶 where id="&info(9),conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('對不起，你不是江湖中人');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
else
mymoney=rs("銀兩")
if rs("魅力")<1 or rs("銀兩")>100 and rs("狀態")<>"獄" and rs("狀態")<>"死" then
randomize timer
r=int(rnd*3)
if r=1 then
conn.execute "update 用戶 set 銀兩=銀兩-100,魅力=魅力+12 where id="&info(9)
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "經過美容院小姐的一番擺弄,你的魅力增加了12點,美容院小姐收取了你100兩銀子!"
response.end
else
conn.execute "update 用戶 set 銀兩=銀兩-10 where id="&info(9)
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
response.write "美容院的學徒由於學藝不精,弄了半天也沒有幫你增加魅力.白花了你十兩服務費!"
response.end
end if
else
response.write "沒錢也想扮靚?美容院小姐們一番譏諷把你轟走門口"
rs.close
set rs=nothing
conn.close
set conn=nothing
response.end
end if
%>
<%end if%>
