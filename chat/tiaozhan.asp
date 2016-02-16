<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
yn=LCase(trim(request.querystring("yn")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名,門派,名單頭像 FROM 用戶 WHERE id="&id,conn
peiou=rs("姓名")
menpai=rs("門派")
old=rs("名單頭像")
if info(0)<>trim(Application("dantiao")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你想作什麼，沒人向你下挑戰書啊！');}</script>"
	Response.End
end if
if yn=1 then
	Response.Write "<script language=JavaScript>{alert('我答應你的要求了！');}</script>"
	conn.execute "update 用戶 set 保護=false where id="&info(9)
	conn.execute "update 用戶 set 保護=false where id="&id
	diantiao="同意單挑：<bgsound src=wav/pk.wav loop=1>["& info(0) &"]與{"& peiou &"}達成協議、開始單挑，方法：用鼠標打擊其頭像！！如已單挑而不單挑，2分鐘內不打,官府將取消單挑！！<br><marquee width=100% behavior=alternate scrollamount=5><font color=FFD060 size=2>★   "& info(0) &"<a href='tiaozhan1.asp?id=" & info(9)&"' target='d'><img src='../ico/"& info(10) &"-2.gif' width='24' height='24' border='0'></a>["&info(5)&"]---------"& peiou &"<a href='tiaozhan1.asp?id=" & info(9)&"' target='d'><img src='../ico/"& old &"-2.gif' width='24' height='24' border='0'></a>["&menpai&"]    開始比武單挑！★</font></marquee>"
rs.close
rs.open "select top 1 姓名2,姓名1 FROM 單挑 WHERE 姓名2='無' or 姓名1='無'",conn
'conn.execute "insert into 單挑(姓名2) values ('"&info(0) &"')"
conn.execute "update 單挑 set 姓名2='"&info(0)&"'" 
conn.execute "update 單挑 set 姓名1='"&peiou&"'"
else
	Response.Write "<script language=JavaScript>{alert('我才不干呢！');}</script>"
	conn.execute "update 用戶 set 魅力=魅力-1000 where id="&id
	diantiao="【拒絕單挑】"&info(0)&"對"&peiou&"說：算了，你厲害我服了你，下次等我武功練成再來與你決一死戰  !{"& info(0) &"}向["& peiou &"]拒絕單挑，魅力下將1000點!"
rs.close
rs.open "select top 1 姓名1,姓名2 FROM 單挑 ",conn
conn.execute "update 單挑 set 姓名1='無',姓名2='無'"
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
sd(195)=info(0)
sd(196)="FFD060"
sd(197)="FFD060"
sd(198)="對"
sd(199)="<font color=FFD060>【江湖消息】</font>"&diantiao
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application("dantiao")=""
Application.UnLock
%>
