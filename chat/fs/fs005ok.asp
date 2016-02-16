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
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以使用迷魂大法！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,門派,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<9 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=9-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("法力")<3000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得3000點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 法力,門派 FROM 用戶 WHERE 姓名='"&to1&"'",conn
falit=int(rs("法力")/2)
if Application("ljjh_dubozhang2")=to1 then
Response.Write "<script language=JavaScript>{alert('對方現在是法力賭局的莊家不可以吸其法力！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 法力=法力-2000+"&falit&",操作時間=now() where id="&info(9)
conn.execute "update 用戶 set 法力=法力-"&falit&" where 姓名='"&to1&"'"
mihunfa="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>對<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>施展<bgsound src=wav/fx40a.wav loop=1>迷魂大法，<font color=FFC01F>["&to1&"]</font>法力逐漸被<font color=red>["&info(0)&"]</font>所吸收掉"&falit&"點，<font color=FFC01F>["&to1&"]</font>獃獃地在那傻笑著，呼..." 
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
sd(194)=to1
sd(195)=info(0)
sd(199)="<font color=CFE0B0>【迷魂術】"& mihunfa &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>