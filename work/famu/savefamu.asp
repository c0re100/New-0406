<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
my=info(0)
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"work/famu")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
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
	Response.Write "<script Language=Javascript>alert('提示：您的職業有錯，所以您不能在這裡進行伐木！！');</script>"
	response.end
end if
if Session("minets")>0 then
	Session("minets")=Session("minets")-1
end if
if Session("minets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你一根木頭也沒有，你拿什麼存？？？！');</script>"
	response.end
end if
rs.close
rs.open "select id from 物品 where 物品名='木頭' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('木頭',"&info(9)&",'物品',0,0,100,2015,2015,0,0,0,"& Session("minets") &",800,"&aaao&")"
else
	id=rs("id")
	conn.execute  "update 物品 set 銀兩=800,數量=數量+" & Session("minets") & " where id="&id
end if
rs.close
rs.open "select 數量 from 物品 where 物品名='木頭' and 擁有者="&info(9)
Session("minesl")=rs("數量")
Response.Write "<script Language=Javascript>alert('您砍到木頭："& session("minets") &"根,現擁有木頭"& rs("數量") &"根，多多努力！');</script>"
Session("minets")=0
Session("cbsj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
