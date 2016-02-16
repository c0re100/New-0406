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
rs.open "select 法力,銀兩,等級,道德,道德加 FROM 用戶 WHERE id="&info(9),conn
if rs("銀兩")<1000000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你身上沒有100萬，先帶100萬在身上再說吧！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
if rs("法力")<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得1000點啊！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
conn.execute "update 用戶 set 法力=法力-1000,道德=道德+10000,銀兩=銀兩-1000000 where id="&info(9)
conn.execute "update 用戶 set 銀兩=銀兩+1000000 where 姓名='"&to1&"'"
rs.close
rs.open "select 等級,道德,道德加 FROM 用戶 WHERE id="&info(9),conn
ddj=(rs("等級")*1440+2200)+rs("道德加")
if rs("道德")>=ddj then
conn.execute "update 用戶 set 道德=道德+"&ddj&" where id="&info(9)
end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
bushishu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>運用法力一式布施天下發給<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>銀兩<font color=red>100萬</font>兩，自己道德增加<font color=red>10000</font>點." 

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
sd(199)="<font color=CFE0B0>【布施術】"& bushishu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>