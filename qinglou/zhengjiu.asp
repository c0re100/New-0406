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
id=request("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select 姓名,美貌度 from 名妓 where ID=" & id
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('這裡沒有這位姑娘呀？！');location.href = 'xiaojie.asp';}</script>"
	response.end
end if
username=rs("姓名")
meimao=rs("美貌度")
rs.close
rs.open "SELECT 內力,武功,操作時間 FROM 用戶 WHERE id="&info(9)
sj=DateDiff("n",rs("操作時間"),now())
if sj<2 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=2-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]分,再操作！');location.href = 'xiaojie.asp';}</script>"
	Response.End
end if
if rs("內力")<0 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你沒內力了呀？！');location.href = 'xiaojie.asp';}</script>"
	response.end
end if
if rs("內力")<10000 and rs("武功")<10000 then
mess=""&info(0)&"武功差勁,想救姑娘["&username&"]沒救出,自己反而被怡紅院的護衛毒打了一頓,魅力下降"&meimao&"點"
conn.execute "update 用戶 set 魅力=魅力-'"&meimao&"' where id="&info(9)
else
mess=""&info(0)&"武功高強、成功救出一位姑娘["&username&"],道德上升"&meimao&"點!"
randomize timer
r=int(rnd*5)+1
if r=4 then
rs.close
rs.open "select id from 物品 where 物品名='天仙秘籍' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,會員) values ('天仙秘籍',"&info(9)&",'物品',0,0,0,0,0,1,200000,99009,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if
mess=""&info(0)&"武功高強、成功救出一位姑娘["&username&"],道德上升"&meimao&"點,官府獎勵《天仙秘籍》一本<img src='../hcjs/jhjs/001/99009.gif' border='0'>"
end if
conn.execute "update 用戶 set 內力=內力-10000,道德=道德+5000,操作時間=now() where id="&info(9)
conn.execute "delete from 名妓 where 姓名='"& username &"'"

end if
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
sd(195)="大家"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【拯救行動】"&info(0)&"</font>在怡紅院拯救無辜少女：<font color=FFD7F4>"&mess&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('拯救行動結束！');location.href = 'xiaojie.asp';}</script>"
	response.end
%> 