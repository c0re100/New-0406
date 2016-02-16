<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
if session("nowinroom")<>1 then
Response.Write "<script Language=Javascript>alert('你在幫派怪獸戰區嗎？');Javascript:window.close();</script>"
	Response.End
end if
pkshijian=Hour(now())
if pkshijian<>21 then
Response.Write "<script Language=Javascript>alert('看好時間了，幫派戰從晚20：00到21：00，現在還早！');Javascript:window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * from 幫派大戰 where 主戰幫派='"&info(5)&"' or 敵對幫派='"&info(5)&"'",conn
if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
  response.write "<Script>alert('[人家幫派大戰你瞎攪合什麼？！');Javascript:window.close();</script>"
  response.end
end if
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
onliners=Application("ljjh_chatrs"&session("nowinroom"))
online=Split(Trim(useronlinename)," ",-1)
x=UBound(online)
if x>0 then
rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
  response.write "<Script>alert('[比賽規則：隻能有一個勝利者，獎品:龍珠一個由該派幫主所有，其門下弟子銀兩1000萬、金幣1000、法力1000！');Javascript:window.close();</script>"
  response.end
end if
name=online(0)
rs.close
	rs.open "select 門派 from 用戶 where 姓名='"&name&"'",conn
menpai=rs("門派")
rs.close
	rs.open "select 掌門 from 門派 where 門派='"&menpai&"'",conn
zhangmen=rs("掌門")
rs.close
	rs.open "select id from 用戶 where 姓名='"&zhangmen&"'",conn
idd=rs("id")
rs.close
rs.open "select id from 物品 where 物品名='龍珠' and 擁有者=" & idd,conn
			if rs.eof and rs.bof then
			conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,堅固度,內力,體力,防御,數量,銀兩,說明,等級,會員) values ('龍珠',"&idd&",'物品',0,100,0,0,0,1,10000000,50000,0,"&aaao&")"			
                        else 
id=rs("id")
conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
 end if
conn.Execute"update 用戶 set 存款=存款+10000000,金幣=金幣+1000,法力=法力+1000 where 門派='" & menpai & "'"
conn.Execute"update 門派 set 掌門='"&name&"' where 門派='" & menpai & "'"
call boot(info(0))
bangpaizhan="恭喜["&info(5)&"]獲得勝利!"
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)=info(0)
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【幫派大戰】</font>"&bangpaizhan
sd(200)=0
Application("ljjh_sd")=sd
Application.UnLock
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
  response.write "<Script>alert('[恭喜你派獲得勝利!基於安全問題請重新進入！]');Javascript:window.close();</script>"
  response.end
sub boot(to1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
%>
 