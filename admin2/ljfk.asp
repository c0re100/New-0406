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
username=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩,操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
if rs("銀兩")<2000000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你沒有200萬怎麼玩遊戲啊！');window.close();}</script>"
Response.End
end if
conn.execute "update 用戶 set 銀兩=銀兩-5000000,操作時間=now() where id="&info(9)
rs.close
rs.open "select id from 物品 where 物品名='獨孤秘籍' and 擁有者=" & info(9),conn
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,會員) values ('獨孤秘籍',"&info(9)&",'物品',0,0,0,0,0,1,200000,990010,"&aaao&")"

	else
id=rs("id")
		conn.execute "update 物品 set 數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者="&info(9)

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
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=FFCFCF>【小道消息】</font><font color=red>"&info(0)&"</font>花了200萬兩銀子開始打反恐精英遊戲了,獲得一本《獨孤秘籍》<img src='../hcjs/jhjs/001/990010.gif' border='0'>！！！"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>

<head>
<title>反恐精英</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false background="../images/8.jpg" leftmargin="0" topmargin="0">
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="550" height="400">
  <param name=movie value="cs.swf">
  <param name=quality value=high>
  <embed src="cs.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="550" height="400">
  </embed> 
</object> 
</body>
