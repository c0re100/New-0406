<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩 from 用戶 where id="&info(9),conn
if rs("銀兩")<1000000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你沒有100萬怎麼玩遊戲啊！');window.close();}</script>"
Response.End
end if
conn.execute "update 用戶 set 銀兩=銀兩-1000000 where id="&info(9)
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
sd(199)="<font color=FFCFCF>【小道消息】</font><font color=red>"&info(0)&"</font>花了100萬兩銀子開始玩麻將了！！！"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>

<head>
<title>靈劍麻將</title>
</head>

<body bgcolor="#000000" text="#FFFFFF" oncontextmenu=self.event.returnValue=false background="../images/8.jpg" leftmargin="0" topmargin="0">
<object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="450" height="338">
  <param name=movie value="bjmj.swf">
  <param name=quality value=high>
  <param name="LOOP" value="false">
  <embed src="bjmj.swf" quality=high pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="450" height="338" loop="false">
  </embed> 
</object>
</body>
