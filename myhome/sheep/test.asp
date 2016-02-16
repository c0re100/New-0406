<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="data.asp"-->
<%
info=Session("info")
if info(0)="" then response.redirect"../../error.asp?id=210"
sql="SELECT * FROM 寵物狀態 where user='"&info(0)&"' and 名字='"&request("name")&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('您的寵物已經死了或您不是這隻寵物的主人！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
name=request("name")
say=rs("說明")
if rs("識別")>=rs("行動點") then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的寵物累了該歇歇了！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
randomize
rad=Int(62 * Rnd + 20)
if rad>0 and rad<30 then rad=1
if rad>30 and rad<60 then rad=2
if rad>60 and rad<90 then rad=3
Select Case rad
case "1"
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,攻擊=攻擊+100,防御=防御+50 where user='"&info(0)&"'"
conn.Execute(sql)
mess="經過了鍛練，你的寵物增加了攻擊100點，防御50點。"
conn.close
case "2"
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,攻擊=攻擊+300,防御=防御+100 where user='"&info(0)&"'"
conn.Execute(sql)
mess="經過了鍛練，你的寵物增加了攻擊300點，防御100點。"
conn.close
case "3"
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,攻擊=攻擊+500,防御=防御+200 where user='"&info(0)&"'"
conn.Execute(sql)
mess="經過了鍛練，你的寵物增加了攻擊500點，防御200點。"
conn.close
Case else
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,攻擊=攻擊+150,防御=防御+100 where user='"&info(0)&"'"
conn.Execute(sql)
mess="經過了鍛練，你的寵物增加了攻擊150點，防御100點。"
conn.close
End Select
abc="<a href='dacu.asp?name="&info(0)&"' target='d'><img src='../myhome/sheep/image/"&say&".gif' border='0' width='94' height='54'></a>"
msg=info(0)&"<font color=red>的寵物["&name&"]來江湖裡鍛煉了,大家快打死它，哈哈!</font><br><marquee width=100%  scrollamount=4>"&abc&"</marquee>"
Application.Lock
Application("ljjh_bianfu")=1
Application.UnLock
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>"&msg&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>鍛練</title>
<link rel="stylesheet" href="setup.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000" background="../../linjianww/0064.jpg">
<div align="center">
<center>
<br>
<br>
</center>
</div>
<div align="center">
<center>
<table border="1" width="378" cellspacing="1" cellpadding="0" height="173" bordercolor="#000000">
<tr>
<td class="p2" width="100%">
<div align="center"><font size="2" color="#000000"> <%=mess%><br>
<br>
</font> </div>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">
<td class="p3" width="100%" height="19">
<p align="center"><font color="#000000"><a href="feedsheep.asp">&gt;&gt;返回</a>
</font>
</td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>

</html>