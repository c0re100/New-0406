<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
id=lcase(trim(request("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 用戶 where id="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你不是江湖中人，滾！');location.href = 'card.asp';}</script>"
	Response.End
end if
rs.close
rs.open "SELECT 物品名,銀兩,說明,攻擊,防御,體力,內力,堅固度,等級,sm FROM 物品買 where ID=" & id & " and 類型='卡片'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('沒有你要購買的會員卡片？!');location.href = 'card.asp';}</script>"
	response.end
end if
card=rs("物品名")
hyyin=rs("銀兩")
hyyin1=rs("攻擊")
say=rs("說明")
sm=rs("sm")
fy=rs("防御")
tl=rs("體力")
nl=rs("內力")
jgd=rs("堅固度")
dj=rs("等級")
if hyyin>=1 then
rs.close
rs.open "select 會員費 from 用戶 where id="&info(9)
if hyyin>rs("會員費") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的會員費不足！!');location.href = 'card.asp';}</script>"
	response.end
end if
conn.execute "update 用戶 set 會員費=會員費-" &hyyin & " where id="&info(9)
rs.close
rs.open "select 會員費 from 用戶 where id="&info(9)
if rs("會員費")<0 then
conn.execute "update 用戶 set 會員費=0 where id="&info(9)
end if
else
rs.close
rs.open "select 金幣 from 用戶 where id="&info(9)
if hyyin1>rs("金幣") then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的金幣不足！!');location.href = 'card.asp';}</script>"
	response.end
end if
conn.execute "update 用戶 set 金幣=金幣-" &hyyin1 & " where id="&info(9)
rs.close
rs.open "select 金幣 from 用戶 where id="&info(9)
if rs("金幣")<0 then
conn.execute "update 用戶 set 金幣=0 where id="&info(9)
end if
end if
rs.close
rs.open "select id,數量 from 物品 where 類型='卡片' and 物品名='" & card & "' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品 (物品名,擁有者,類型,攻擊,防御,內力,體力,堅固度,銀兩,說明,數量,sm,等級,會員) values ('"&card&"',"&info(9)&",'卡片',"&hyyin1&","&fy&","&nl&","&tl&","&jgd&","&hyyin&",'"&say&"',1,"&sm&","&dj&","&aaao&")"
else
	id=rs("id")
	sl=rs("數量")
	conn.execute "update 物品 set 銀兩=" & hyyin & ",數量=數量+1,會員="&aaao&" where id="&id
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的卡片["&card&"]購買成功,現有"&sl+1&"個！');location.href = 'card.asp';}</script>"
response.end
%>
