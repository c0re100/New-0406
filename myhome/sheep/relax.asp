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
if rad>0 and rad<20 then rad=1
if rad>20 and rad<40 then rad=2
if rad>40 and rad<60 then rad=3
if rad>60 and rad<80 then rad=4
Select Case rad
case "1"
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,體力=體力+50 where user='"&info(0)&"'"
conn.Execute(sql)
mess="寵物得到休息，增加體力50點。"
conn.close
case "2"
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,體力=體力+100 where user='"&info(0)&"'"
conn.Execute(sql)
mess="寵物得到休息，增加體力100點。"
conn.close
case "3"
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,體力=體力+200 where user='"&info(0)&"'"
conn.Execute(sql)
mess="寵物得到休息，增加體力200點。"
conn.close
case "4"
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,體力=體力+150 where user='"&info(0)&"'"
conn.Execute(sql)
mess="寵物得到休息，增加體力150點。"
conn.close
Case else
name=request("name")
sql="update 寵物狀態 set 識別=識別+1,體力=體力+10 where user='"&info(0)&"'"
conn.Execute(sql)
mess="寵物得到休息，增加體力10點。 "
conn.close
End Select
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>體息</title>
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