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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(4)=0 then 
aaao=0
else
aaao=1
end if
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
if fla<100 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得100點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.End
end if
sj=DateDiff("s",rs("操作時間"),now())
if sj<15 then
rs.close
set rs=nothing	
conn.close
set conn=nothing
ss=15-sj
Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
conn.execute "update 用戶 set 法力=法力-100,操作時間=now() where id="&info(9)
randomize()
rnd1=int(rnd()*5)
if rnd1<3 then
banbianshu=info(0) & "好可惜，你尋遍了江湖各個角落也沒有找到什麼水晶球,"&info(0) & "損耗100點法力!" 
else
rs.close
rs.open "select id FROM 物品 WHERE 物品名='水晶球' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明,sm,等級,會員) values ('水晶球',"&info(9)&",'法寶',100,0,0,0,0,1,200000,1600,1600,0,"&aaao&")"
	else
id=rs("id")
		conn.execute "update 物品 set 銀兩=200000,數量=數量+1,會員="&aaao&" where id="& id &" and 擁有者=" & info(9)
	end if
banbianshu=info(0)& "你千辛萬苦尋訪<font color=red>水晶球</font>的下落，卻不料在一條水溝裡被你找到了，趕緊洗洗水晶球挖掘它的魔力吧." 
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
sd(194)=info(9)
sd(195)=info(0)
sd(199)="<font color=CFE0B0>【尋找水晶】"& banbianshu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>