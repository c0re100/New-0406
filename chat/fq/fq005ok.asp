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
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以使用絕情刀！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
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
if fla<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得1000點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 法力=法力-1000,操作時間=now() where id="&info(9)
rs.close
rs.open "select 體力,門派 FROM 用戶 WHERE 姓名='"&to1&"'",conn
tl=rs("體力")
money=int(rs("體力")/3)
rs.close
rs.open "select id FROM 物品 WHERE 物品名='絕情刀' and 數量>=0 and 擁有者=" & info(9)
if rs.eof or rs.bof then
Response.Write "<script language=JavaScript>{alert('你有絕情刀嗎？');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
id=rs("id")
conn.execute "update 物品 set 數量=數量-1 where id="& id &" and 擁有者=" & info(9)
r=int(rnd*2)+1
select case r
case 2
conn.execute "update 用戶 set 體力=體力-"&money&" where 姓名='"&to1&"'"
jiuqingdao="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>從腰間撥出一把絕情刀<bgsound src=wav/fx42.wav loop=1>，<img src='img/9.gif' width='44' height='44'>惡狠狠地對著<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>刺了下去,受死吧，"&towho&"啊的一聲，被刺中心脈,曾經滄海難為水啊，體力失去1/3，共計"&money&"點......" 
if tl<=0 then 
conn.execute "update 用戶 set 狀態='死' where 姓名='"&to1&"'"
jiuqingdao="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>從腰間撥出一把絕情刀<bgsound src=wav/ZR2199.wav loop=1>，<img src='img/9.gif' width='44' height='44'>惡狠狠地對著<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>刺了下去,受死吧,"&towho&"啊的一聲，由於體力不支，被當場刺死，自古多情空餘恨啊..." 
 call boot(to1)
end if
case else
jiuqingdao="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>從腰間撥出一把絕情刀<bgsound src=wav/fx42.wav loop=1>，<img src='img/9.gif' width='44' height='44'>惡狠狠地對著<a href=javascript:parent.sw('[" & to1 & "]'); target=f2><font color=red>["&to1&"]</font></a>刺了下去,受死吧,怎料是愛到深處心恨誰，沒刺到啊......"
end select
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
sd(199)="<font color=FFC01F>【絕情刀】"&jiuqingdao&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
'踢人
sub boot(to1)
Application.Lock
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
dim newonlinelist()
useronlinename=""
onliners=0
js=1
for i=1 to UBound(onlinelist) step 6
if CStr(onlinelist(i+1))<>CStr(to1) then
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
else
kickip=onlinelist(i+2)
end if
next
useronlinename=useronlinename&" "
if kickip<>"" then
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
else
Application.UnLock
Response.Redirect "manerr.asp?id=204&kickname=" & server.URLEncode(kickname)
end if
Application.UnLock
end sub
%>