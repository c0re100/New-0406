<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
	Response.Write "<script Language=Javascript>alert('所攻擊的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 等級,法力,身份,門派,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<6 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=6-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
dj=rs("等級")
shenfen=rs("身份")
fla=rs("法力")
if fla<10000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得100點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
if dj>100 or shenfen="掌門" or rs("門派")="官府" then
Response.Write "<script language=JavaScript>{alert('你是掌門啊或等級100級了這麼厲害還用嫵媚術呀！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 性別,銀兩 FROM 用戶 WHERE 姓名='"&to1&"'",conn
xb=rs("性別")
money=int(rs("銀兩")/4)
randomize timer
kaa=int(rnd*99)+1
if money>=1000000 then
randomize timer
moneyc=1000000
moneyc=killer+kaa
end if
if xb="女" then
Response.Write "<script language=JavaScript>{alert('此嫵媚術只對男孩子適用！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 銀兩=銀兩+"&moneyc&",法力=法力-10000,操作時間=now() WHERE  id="&info(9)
conn.execute "update 用戶 set 銀兩=銀兩-"&money&" where 姓名='"&to1&"'"

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
meihuoren="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>對<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>施展<bgsound src=wav/phant030a.wav loop=1>魅惑人間的媚術，<font color=FFC01F>"&to1&"</font>被<img src='img/007.gif'><font color=red>["&info(0)&"]</font>傾國傾城的容貌所迷惑，身上的銀子不知不覺少了四分之一,共計<font color=red>"&moneyc&"</font>兩:)..." 
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
sd(194)=to1
sd(195)=info(0)
sd(199)="<font color=CFE0B0>【嫵媚術】"& meihuoren &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>