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
rs.open "select 職業 from 用戶 where id="&info(9),conn

if Session("tiets")<=1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你一塊鐵也沒有，你拿什麼賣錢？？？！');</script>"
	response.end
end if
if Session("tiesl")<=0 then
	Session("tiesl")=0
end if
conn.execute "update 用戶 set 銀兩=銀兩+"& Session("tiets")*3000 &" where id="&info(9)
conn.execute "update 物品 set 數量="& Session("tiesl") &" where 物品名='礦石' and 擁有者="&info(9)
Response.Write "<script Language=Javascript>alert('您采鐵塊："& session("tiets") &"塊,賣了"& session("tiets")*3000 &"兩白銀！');</script>"
Session("tiets")=0
Session("tiesj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>