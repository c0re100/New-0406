<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")%>
<html>
<head>
<title>江湖消息</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#88AFD7">
<%
my=info(0)
if my<>"" then
Response.Write "<div align='center'><font size=-1>師傅："& info(7) &"</font></div>"
rs.open "select 師傅,師傅交錢,銀兩,配偶,結婚時間,性別,小孩 from 用戶 where 姓名='" & my & "'",conn
if rs("師傅")<>"" and rs("師傅")<>"無" and rs("師傅交錢")<>"交錢"&Day(date()) then
	yin=int(rs("銀兩")*0.05)
	conn.execute "update 用戶 set 銀兩=銀兩-"& yin &",師傅交錢='交錢"& Day(date()) &"' where 姓名='"& my &"'"
	conn.execute "update 用戶 set 銀兩=銀兩+"& yin &" where 姓名='"& info(7) &"'"
	Response.Write "<div align='center'><font size=-1>"& my &"今天上交師門"& yin &"兩白銀，來孝敬師傅："& info(7) &"</font></div>"
end if
if rs("配偶")<>"無" and rs("性別")<>"男" then
	Response.Write "<br>你在江湖已經結婚：結婚日期為："& rs("結婚時間") &"，注意：20天後分娩！"
end if
if rs("配偶")<>"無" and date()-rs("結婚時間")>19 and rs("性別")<>"男" and rs("小孩")="無" then 
        Response.Write "<br>20天到了，<input type=button value='要孩子了' onClick=window.open('haizi.asp?id="&info(9)&"&yhz=1')><input type=button value='不要孩子' onClick=window.open('haizi.asp?id="&info(9)&"&yhz=0')>"
end if
rs.close
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r
rs.open "select 貸款日期,貸款人 from 貸款 where 還貸記錄=false and 貸款人='"& my & "'",conn
if not(rs.BOF or rs.EOF) then
	Response.Write "<br>你在江湖錢裝有貸款，貸款日期："& rs("貸款日期") &"注意還款,7天不還刪除ID！"
end if
rs.close
rs.open "select * from 貸款 where 還貸記錄=false and 貸款人='"& my & "' and DateDiff('d',貸款日期,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	name=rs("貸款人")
	conn.Execute ("update 貸款 set 還貸記錄=true where 貸款人='"&name&"'")
	conn.Execute ("delete * from 用戶 where 姓名='"&name&"'")
	'conn.Execute ("insert into 人命(死者,時間,兇手,死因) values ('" & name & "',"&sqlstr(sj)&",'高利貸殺手','欠債還錢，沒錢要你命')")
	Response.Write "因為你7天沒有還貸款，江湖ID被刪除了！"
	info(0)=""
	Session.Abandon
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end if
else%>
<font size=-1>靈劍江湖歡迎您！</font> 
<%end if
%>
<div align="center"></div>
</body>
</html>
 