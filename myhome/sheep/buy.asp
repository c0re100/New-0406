<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%><!--#include file="data.asp"--><%
Response.Expires=0
info=Session("info")
if info(0)="" then
response.redirect"warning.htm"
else
name=request("name")
rs.open"select 價格,攻擊,防御,體力,寵物類型,技能,說明,id from 寵物 where 寵物類型='"&name&"'",conn,1,1
money1=rs("價格")
at=rs("攻擊")
guard=rs("防御")
power=rs("體力")
lx=rs("寵物類型")
jn=rs("技能")
say=rs("說明")
petid=rs("id")
rs.close
rs.open"select 銀兩 from 用戶 where 姓名='"&info(0)&"'",conn,1,1
if rs("銀兩")-money1<0 then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"您沒有足夠的錢買寵物！",0,"FLUSH"
history.back
</script>
<%
else
money=rs("銀兩")-money1
conn.execute("update 用戶 set 銀兩='"&money&"'  where 姓名='"&info(0)&"'")
rs.close%>
<!--#include file="data2.asp"-->

<%rs.open"select 名字 from 寵物狀態 where user='"&info(0)&"'",conn,1,1
if not rs.bof then
rs.close
conn.close
%>
<script language="vbscript">
msgbox"一位大俠隻能買一隻寵物呵",0,"FLUSH"
history.back
</script>
<%
else
set rs=server.createobject("adodb.recordset")
sql="select 名字,說明,寵物類型,技能,攻擊,體力,價格,防御,類型id,user from 寵物狀態 where (id is null)"
rs.open sql,conn,1,3
rs.addnew
rs("名字")=lx
rs("說明")=say
rs("寵物類型")=lx
rs("技能")=jn
rs("攻擊")=at
rs("體力")=power
rs("價格")=money1
rs("防御")=guard
rs("類型id")=petid
rs("user")=info(0)
rs.update
rs.close
set rs=nothing	
conn.close
set conn=nothing
end if
end if
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
<p align="center"><font size="2">您已經成功購買了<%=lx%> ，快回去幫寵物起個好聽的名字！<br>
<br>
</font>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">

<td class="p3" width="100%" height="19">
                <p align="right"><font size="2"><a href="feedsheep.asp" target="_top"><font color="#00FFFF">快開始培養<%=lx%>吧！</font></a></font> 
              </td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>

</html>
