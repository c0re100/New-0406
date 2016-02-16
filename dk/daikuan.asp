<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select  allvalue,銀兩,等級,存款 from 用戶 where id="&info(9),conn
allvalue=rs("allvalue")
bigdk=allvalue*100
yinliang=rs("銀兩")
jhdenji=rs("等級")
nowyl=rs("銀兩")
nowck=rs("存款")
rs.close
rs.open "select 貸款人 from 貸款 where 還貸記錄=false and DateDiff('d',貸款日期,#" & sj & "#)>7",conn
if not(rs.BOF or rs.EOF) then
	do while not (rs.bof or rs.eof)
		name=rs("貸款人")
		conn.Execute ("update 貸款 set 還貸記錄=true where 貸款人='"&name&"'")
		'conn.Execute ("delete * from 用戶 where 姓名='"& name &"'")
		conn.Execute ("insert into 人命(死者,時間,兇手,死因) values ('" & name & "',"&sqlstr(sj)&",'高利貸殺手','欠債還錢，沒錢要你命')")
		rs.movenext
		name=""
		info(0)=""
	loop
end if
%>
<html>
<head>
<title>貸款</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="chat/READONLY/Style.css">
</head>

<body bgcolor="#FFFFFF" background="../images/8.jpg">
<p align="center"><font size="6" face="隸書">江湖高利貸款</font></p>
<p align="center">&nbsp;</p>
<form method="post" action="dk.asp">
<table width="300" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#E0E0E0">
<tr>
<td>申貸人:</td>
<td><%=info(0)%></td>
</tr>
<tr>
<td>現在銀兩：</td>
<td><%=nowyl%>兩</td>
</tr>
<tr>
<td>現在存款:</td>
<td><%=nowck%>兩</td>
</tr>
<tr>
<td>最大貸款金額:</td>
<td><%=bigdk%>兩</td>
</tr>
<tr>
<td>貸款金額：</td>
<%
rs.close
rs.open "select 貸款金額,貸款日期 from 貸款 where 貸款人='"&info(0)&"' and 還貸記錄=false",conn
if rs.BOF or rs.EOF then
	%>
	<td>
	<%if jhdenji>3 then%>
		<input type="text" name="dk" size="10" maxlength="10">
	<%else%>
		<font color=red>不能操作</font>
	<%end if%>
	</td>
	</tr>
<tr>
<td colspan="2">
<div align="center">
<%if jhdenji>3 then%>
	<input type=submit value="貸款" name="submit">
	<input type="reset" name="reset" value="清空">
<%else%>
	<font color=red>戰鬥小於3級江湖不貸款給你！</font>
<%end if%>
</div>
</td>
<%else%>
<td>
<%if yinliang>rs("貸款金額")*1.5 then%>
<input type="text" name="dk" size="10" maxlength="10" value='<%=rs("貸款金額")%>' readonly>
<%else%>
<font color=red>不能操作</font>
<%end if%>
<br>貸款日期：<%=rs("貸款日期")%>

</td>
</tr>
<tr>
<td colspan="2">
<div align="center">
<%if yinliang>rs("貸款金額")*1.5 then%>
<input type=submit value="還款" name="submit">
<input type="reset" name="reset" value="清空">
<%else%>
<font color=red>你的現金不夠還款!</font>
<%end if%>
</div>
</td>
<%end if
rs.close
set rs=nothing
conn.Close
set conn=nothing%>
</tr>
</table>
</form>
<p align="center">本錢莊提供貸款，小人物不貸（戰鬥等級不到3級不放貸）<br>
貸款期限是一個星期7天。俗話說&quot;欠債還錢，沒錢還命&quot;，到期不還者本莊將雇<br>
殺人將其殺之，<font color=red>(將刪除江湖ID)</font>望各位大俠相互轉告！！！！！！ <br>
(高利貸還錢比例為貸100兩，還150兩，不是我心黑呀，現在賺錢難呀！) </p>
</body>
</html>
<%
Function SqlStr(data)
SqlStr="'" & Replace(data,"'","''") & "'"
End Function
%>