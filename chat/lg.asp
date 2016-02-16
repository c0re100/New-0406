<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
Response.Buffer=true
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(2)=0 then
Response.Write "<script Language=Javascript>alert('怪物房是不能開啟保護的！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM 通緝令 WHERE 通緝人犯='" & info(0) & "' and 站長審批='批準'",conn
if not(rs.bof or rs.eof)  then	
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是在逃通輯犯，不能開保護！');}</script>"
	Response.End
end if
rs.close
rs.open "select top 1 姓名1,姓名2 FROM 單挑 ",conn
if rs("姓名1")=info(0) or rs("姓名2")=info(0)  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你現在在單挑狀態，不得逃跑！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
rs.close
rs.open "SELECT 殺人數,寶物,保護,操作時間,會員等級 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
aaay=rs("會員等級")
if sj<60 and aaay=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if

if rs("殺人數")>=30 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('作盡天下壞事,還想保護？！');}</script>"
	Response.End
end if
if rs("寶物")="靈劍水晶石" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你現在有江湖寶物[靈劍水晶石]不能進行練功保護！');}</script>"
	Response.End
end if
if rs("保護")=true then
	conn.Execute "update 用戶 set 保護=false,操作時間=now() where id="&info(9)	
	diaox="功成出關，300年前的武林慘案又將再次重演了！"
else
	conn.Execute "update 用戶 set 保護=true,操作時間=now() where id="&info(9)
	diaox="為了修煉無上神功，再次閉關，武林中總算可以太平一陣了！"
end if
conn.close
set rs=nothing
set conn=nothing
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
sd(196)="9FDF70"
sd(197)="9FDF70"
sd(198)="對"
sd(199)="<font color=#9FDF70>【消息】" & info(0) & ""& diaox &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd

%>
<script language=javascript>
parent.f4.location.href='f4.asp';
</script> 