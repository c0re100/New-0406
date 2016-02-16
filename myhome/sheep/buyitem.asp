<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
info=Session("info")
if info(0)="" then Response.Redirect "../../error.asp?id=210"
%><!--#include file="data.asp"--><%
name=request("name")
rs.open"select 價格,攻擊,防御,體力 from 寵物道具 where 道具名字='"&name&"'",conn,1,1
money1=rs("價格")
at=rs("攻擊")
guard=rs("防御")
power=rs("體力")
rs.close
rs.open"select 銀兩,操作時間 from 用戶 where 姓名='"&info(0)&"'",conn,1,1
sj=DateDiff("s",rs("操作時間"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'itemshop.asp';}</script>"
	Response.End
end if
if rs("銀兩")-money1<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"您沒有足夠的錢買這個道具！",0,"FLUSH"
history.back
</script>
<%
else
money=rs("銀兩")-money1
conn.execute("update 用戶 set 銀兩="&money&",操作時間=now()  where 姓名='"&info(0)&"'")
rs.close
%><!--#include file="data2.asp"--><%
set rs=server.createobject("adodb.recordset")
sql="select 道具名字,攻擊,體力,價格,防御,擁有者 from 道具列表 where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("道具名字")=name
rs("攻擊")=at
rs("體力")=power
rs("價格")=money1
rs("防御")=guard
rs("擁有者")=info(0)
rs.update
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>購買成功</title>
<link rel="stylesheet" href="setup.css">
</head>

<body bgcolor="#3a4b91" background="../../chat/bk.jpg" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
<center>
<br>
<br>
</center>
</div>
<div align="center">
<center>
    <table border="1" width="338" cellspacing="1" cellpadding="0" height="104" bordercolor="#FFFF00">
      <tr>
<td class="p2" width="100%">
<p align="center"><font size="2">您已經成功購買了<%=name%>。<br>
<br>
<br>
<br>
</font>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">

<td class="p3" width="100%" height="19">
<p align="center">&gt;&gt;<a href="itemshop.asp"><font color="#FFFFFF">返回道具商店</font></a>
</td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>
</html>