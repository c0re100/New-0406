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
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以使用搶劫令！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力 FROM 用戶 WHERE id="&info(9),conn
if rs("法力")<10000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得10000點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
rs.close
rs.open "select 門派,存款 FROM 用戶 WHERE 姓名='"&to1&"'",conn
money=int(rs("存款")/5)
randomize timer
kaa=int(rnd*99)+1
if money>=1000000 then
randomize timer
moneyc=1000000
moneyc=killer+kaa
end if
rs.open "select id FROM 物品 WHERE 物品名='搶劫令' and 數量>0 and 擁有者="&info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有搶劫令嗎？');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update 物品 set 數量=數量-1 where id="& id &" and 擁有者="&info(9)
conn.execute "update 用戶 set 法力=法力-10000 where id="&info(9)
conn.execute "update 用戶 set 銀兩=銀兩+"&moneyc&",道德=道德-1000 WHERE  id="&info(9)
conn.execute "update 用戶 set 存款=存款-"&money&" where 姓名='"&to1&"'"
qiangjielin="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>拿出法器<bgsound src=wav/Phant012.wav loop=1>搶劫令,對著<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>大吼一聲:搶劫  ,從<font color=FFC01F>["&to1&"]</font>那搶得存款<font color=FFC01F>["&moneyc&"]</font>兩,<font color=FFC01F>["&info(0)&"]</font>消耗法力<font color=FFC01F>3000</font>點" 
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
sd(199)="<font color=FFC01F>【搶劫令】"&qiangjielin&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>