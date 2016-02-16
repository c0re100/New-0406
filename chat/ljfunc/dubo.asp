<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
db=request.querystring("db")
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名 FROM 用戶 WHERE id="&id,conn
peiou=rs("姓名")
if info(0)<>trim(Application("ljjh_kissly")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你想作什麼，人家也沒有說和你賭博！');}</script>"
	Response.End
end if
if db-Application("ljjh_bingwen")=-2 or db-Application("ljjh_bingwen")=1 then
	Response.Write "<script language=JavaScript>{alert('你輸了，倒霉催的！哎~~~');}</script>"
	conn.execute "update 用戶 set 銀兩=銀兩-1000000 where id="&info(9)
	conn.execute "update 用戶 set 銀兩=銀兩+2000000 where id="&id
        hunyin="倒霉：["& info(0) &"]賭博裡輸給了{"& peiou &"}，100萬兩白花花的銀子飛了！"
end if
if db-Application("ljjh_bingwen")=-1 or db-Application("ljjh_bingwen")=2 then
        Response.Write "<script language=JavaScript>{alert('你勝了，哈哈，100萬兩哦~~~');}</script>"
	conn.execute "update 用戶 set 銀兩=銀兩+1000000 where id="&info(9)
	hunyin="恭喜：["& info(0) &"]賭博裡贏了{"& peiou &"}100萬兩銀子！"
end if
if db=Application("ljjh_bingwen") then
	Response.Write "<script language=JavaScript>{alert('平了，重新來吧！');}</script>"
        conn.execute "update 用戶 set 銀兩=銀兩+1000000 where id="&id
	hunyin="哈哈：["& info(0) &"]賭博裡和{"& peiou &"}打平了，偉大的頭腦不謀而合。"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
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
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=red>【江湖消息】</font>"&hunyin
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application("ljjh_kissly")=""
Application("ljjh_bingwen")=""
Application.UnLock
%>
