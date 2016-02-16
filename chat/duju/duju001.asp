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
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(5)<>0 then
Response.Write "<script Language=Javascript>alert('要賭博請先去大廳！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
'f=Minute(time())
'if f>30 then
'	Response.Write "<script language=JavaScript>{alert('現在不能賭博！賭博時間為每小時的前30分鐘,例如：17:00開始17:30結束!!');}</script>"
'	Response.End 
'end if
if Request.Form("銀兩局")="銀兩局作莊" then
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select 存款 FROM 用戶 WHERE id=" & info(9),conn
	if rs("存款")<20000000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你的存款不夠2000萬，不能作莊！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 姓名 FROM 在線賭博 WHERE 身份='莊家'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('現在["&rs("姓名")&"]為莊家,你不能作莊！！');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into 在線賭博(姓名,身份,銀兩,押大小,時間) values ('"&info(0) &"','莊家',0,'莊',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application.Lock
	Application("ljjh_dubozhang")=info(0)
Application.UnLock
	dubo="<font color=green>【銀兩賭局作莊】</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>在靈劍江湖在線賭場申請做莊成功，現在大家可以豪賭一把了,快來快來呀  "
 
end if

if Request.Form("法力局")="法力局作莊" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select 法力 FROM 用戶 WHERE id=" & info(9),conn
	if rs("法力")<5000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你的法力不夠5000，不能作莊！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 姓名 FROM 法力賭局 WHERE 身份='莊家'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('現在["&rs("姓名")&"]為莊家,你不能作莊！！');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into 法力賭局(姓名,身份,法力,押大小,時間) values ('"&info(0) &"','莊家',0,'莊',now())"	
	rs.close
Application.Lock
	Application("ljjh_dubozhang2")=info(0)
Application.UnLock
	dubo="<font color=green>【法力賭局作莊】</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>在靈劍江湖法力賭局申請做莊成功，現在大家可以各憑運氣賭一把了,快來快來呀  "
end if

if Request.Form("泡點局")="泡點局作莊" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select allvalue FROM 用戶 WHERE id=" & info(9),conn
	if rs("allvalue")<3000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你的點數不夠3000點，不能坐莊！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 姓名 FROM 點數賭局 WHERE 身份='莊家'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('現在["&rs("姓名")&"]為莊家,你不能坐莊！！');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into 點數賭局(姓名,身份,點數,押大小,時間) values ('"&info(0) &"','莊家',0,'莊',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application.Lock
	Application("ljjh_dubozhang3")=info(0)
Application.UnLock
	dubo="<font color=green>【泡點賭局作莊】</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>在靈劍江湖存點數賭局申請做莊成功，現在大家可以各憑運氣賭一把了,快來快來呀  "
end if

if Request.Form("金幣局")="金幣局作莊" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select 金幣 FROM 用戶 WHERE id=" & info(9),conn
	if rs("金幣")<4000 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你的金幣不夠4000個，不能坐莊！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	rs.open "select top 1 姓名 FROM 金幣賭局 WHERE 身份='莊家'",conn
	if not(rs.eof) or not(rs.bof) then
		Response.Write "<script language=JavaScript>{alert('現在["&rs("姓名")&"]為莊家,你不能坐莊！！');location.href = 'javascript:history.go(-1)';}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End 
	end if
	conn.execute "insert into 金幣賭局(姓名,身份,金幣,押大小,時間) values ('"&info(0) &"','莊家',0,'莊',now())"	
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Application.Lock
	Application("ljjh_dubozhang4")=info(0)
Application.UnLock
	dubo="<font color=green>【金幣賭局作莊】</font><img src='../jhimg/sz.gif'><a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a><bgsound src=wav/Faqian.wav loop=1>在靈劍江湖金幣賭局申請做莊成功，現在大家可以各憑運氣賭一把了,快來快來呀  "
end if

if Request.Form("吆喝")="各賭局吆喝" then
Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 FROM 用戶 WHERE id="&info(9),conn
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
rs.close
tmprs=conn.execute("Select count(*) As 數量 from 在線賭博 where 銀兩<>0")
	durs1=tmprs("數量")
set tmprs=nothing	
	rs.open "select top 1 姓名 FROM 在線賭博 WHERE 身份='莊家' and 姓名='"&info(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你也不是莊，你想搞什麼？！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
tmprs=conn.execute("Select count(*) As 數量 from 法力賭局 where 法力<>0")
	durs2=tmprs("數量")
set tmprs=nothing
	rs.open "select top 1 * FROM 法力賭局 WHERE 身份='莊家' and 姓名='"&info(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你也不是莊，你想搞什麼？！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close

	tmprs=conn.execute("Select count(*) As 數量 from 點數賭局 where 點數<>0")
	durs3=tmprs("數量")
set tmprs=nothing
	rs.open "select top 1 姓名 FROM 點數賭局 WHERE 身份='莊家' and 姓名='"&info(0)&"'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你也不是莊，你想搞什麼？！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
conn.Execute "update 用戶 set 操作時間=now() where id="&info(9)
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	dubo="<font color=green>【賭局吆喝】</font><img src='../jhimg/sz.gif'>莊家:<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>叫到：買得多贏得多，買得少贏得少！買了您後悔(怎麼不多買點呀!)不買您更後悔(要買就中了)，<font color=red>【銀兩賭局】</font>還有:"& (10-durs1) &"個下注位開局  <font color=red>【法力賭局】</font>還有:"& (5-durs2) &"個下注位開局  <font color=red>【泡點賭局】</font>還有:"& (5-durs3) &"個下注位開局  "
end if
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
sd(194)=info(0)
sd(195)=info(0)
sd(199)="<font color=#FFC01F>"& dubo &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{location.href = 'javascript:history.go(-1)';}</script>"
Response.End 
%>