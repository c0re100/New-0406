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
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"choujiang/chou.asp")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
abc=DateDiff("s",Session("choujiang"),now())
if abc<30 or abc>240 then 
	Response.Write "<script Language=Javascript>alert('已經過了領獎時間，請待下次吧！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<1 then
	ss=1-sj
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('請等上"&ss&"分再來踫運氣吧！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
conn.execute "update 用戶 set 操作時間=now()  where id="&info(9)
randomize timer
r=int(rnd*12)+1
'r=9
select case r
case 1
	mess="獲得了大獎，獎品是攻擊、防御、體力各加3億"
	conn.execute "update 用戶 set 攻擊=攻擊+300000000,防御=防御+300000000,體力=體力+300000000 where id="&info(9)

case 2
	mess="獲得了二獎，獎品是經驗5000億"
	conn.execute "update 用戶 set allvalue=allvalue+500000000000 where id="&info(9) 
case 3
	mess="獲得了三獎，獎品是法力10億"
	conn.execute "update 用戶 set 法力=法力+1000000000 where id="&info(9)
case 4
	mess="獲得了四獎，獎品是道德1億"
	conn.execute "update 用戶 set 道德=道德+100000000 where id="&info(9)
case 5
	mess="獲得了五獎，獎品是豆點10萬"
	conn.execute "update 用戶 set 豆點=豆點+100000 where id="&info(9)
case 6
	mess="獲得了六獎，獎品是銀兩10億"
	conn.execute "update 用戶 set 銀兩=銀兩+1000000000 where id="&info(9)
case 7
	mess="獲得了七獎，獎品是天眼"
	conn.execute "update 用戶 set 天眼=1 where id="&info(9)
case 8

        mess="獲得了八獎，獎品是綠名"
	conn.execute "update 用戶 set 性別='無' where id="&info(9)
case 9
	mess="獲得了九獎，獎品是屠龍刀10把"
	rs.open "select id from 物品 where 物品名='屠龍刀' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,sm,會員) values ('屠龍刀',"&idd&",'右手',0,0,0,3400000,900000,20032017,10,300,20032017,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 數量=數量+10,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 10
	mess="獲得了十獎，獎品是至尊劍10把"
	rs.open "select id from 物品 where 物品名='至尊劍' and 擁有者="&info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,防御,內力,體力,堅固度,攻擊,說明,數量,等級,sm,會員) values('至尊劍',"&idd&",'左手',800000,0,0,3200000,0,20032016,10,265,20032016,"&aaao&")"
	else
		id=rs("id")
		conn.execute "update 物品 set 數量=數量+10,會員="&aaao&" where id="& id &" and 擁有者="&info(9)
	end if
	rs.close
case 11
	mess="獲得了特別獎，獎品是會費和金幣各1億"
	conn.execute "update 用戶 set 會員費=會員費+100000000,金幣=金幣+100000000 where id="&info(9)
case else
   mess="獲得了安慰獎，獎品是流星1000個"
rs.open "select id from 物品 where 物品名='流星' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,攻擊,防御,堅固度,說明,sm,等級,內力,體力,數量,銀兩,會員) values ('流星',"&info(9)&",'物品',0,0,1000,111111,111111,0,0,0,1000,1000000,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1000,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

	end if

rs.close
end select


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
sd(199)="<font color=#ff0000>消息</font>"& info(0) &"在抽獎活動中"&mess
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>結果公佈</title></head>

<body oncontextmenu=self.event.returnValue=false background="../jhimg/bk_hc3w.gif" link="#000000" vlink="#000000" alink="#000000">
<div align="center"> <br>
<br>
<table border=1 bgcolor="#948754" align=center cellpadding="10" cellspacing="13" height="186" width="300">
<tr>
<td bgcolor=#C6BD9B>
<table align="center" width="248">
<tr>
<td valign="top">
<div align="center">
<p><%=mess%></p>
</div>
</table>
<div align="center"><br>
<input type=button value=關閉窗口 onClick='window.close()' name="button">
</div>
</tr>
</table>
</div>
</body>
<noembed>
</html>