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
id=request("name")
set rs=conn.execute("select 價格 from 道具列表 where id="&id&"")
if not rs.eof then
money=rs("價格")
money=money/2
rs.close
conn.execute("delete from 道具列表 where id="&id&" and 擁有者='"&info(0)&"'")
conn.close
%>
<!--#include file="data.asp"-->
<%set rs=conn.execute("select 銀兩 from 用戶 where 姓名='"&info(0)&"'")
money1=rs("銀兩")+money
conn.execute("update 用戶 set 銀兩='"&money1&"'  where 姓名='"&info(0)&"'")%>
<link rel="stylesheet" href="setup.css">
<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000" background="../../linjianww/0064.jpg">
<div align="center">你半價賣出了寵物道具得到<%=money%>兩！ <%else%> 你沒有寵物道具不能賣出 <%rs.close%>
<%end if%><br>
<br>
<a href="indexsheep.asp"><font color="#FFFFFF">返回</font></a></div>
