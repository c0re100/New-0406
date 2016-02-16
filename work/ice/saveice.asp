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
if info(4)=0 then 
aaao=0
else
aaao=1
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 職業 from 用戶 where id="&info(9),conn
if trim(rs("職業"))="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的職業不能在這裡進行采冰，請轉換職業！');</script>"
	response.end
end if
if Session("icets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你一桶冰也沒有采，存什麼？');</script>"
	response.end
end if
randomize timer
r=int(rnd*3)
if r=2 then
rs.close
rs.open "select id from 物品 where 物品名='朱石' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('朱石',"&info(9)&",'物品',0,0,0,2003306,2003306,0,0,0,1,10000,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update 物品 set 數量=數量+1,會員="&aaao&" where id="&id
end if
Response.Write "<script Language=Javascript>alert('您采到一塊朱石，此物可以用來煆造絕世武器！');</script>"
end if
rs.close
rs.open "select id from 物品 where 物品名='冰水' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,內力,體力,數量,銀兩,說明,sm,等級,會員) values ('冰水',"&info(9)&",'藥品',0,0,100,151,151,"& Session("icets") &",2000,151,151,0,"&aaao&")"
else
	id=rs("id")
	conn.execute "update 物品 set 銀兩=2000,數量=數量+" & Session("icets") & ",會員="&aaao&" where id="&id
end if
rs.close
rs.open "select 數量 from 物品 where 物品名='冰水' and 擁有者="&info(9),conn
Session("icesl")=rs("數量")
Response.Write "<script Language=Javascript>alert('您采冰："& session("icets") &"桶,現擁有冰水"& rs("數量") &"桶，多多努力！');</script>"
Session("icets")=0
Session("icejl1")=0
Session("cbsj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
