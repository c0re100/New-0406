<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(0)
a=request("a")
id=request("id")
name=ljjh_nickname
name2=request.form("name2")
if a="a" then
sql="SELECT * FROM 用戶 WHERE 姓名='" & name2 & "'"
else
sql="SELECT * FROM 用戶 WHERE ID=" & id
end if
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess=name2 & "不是江湖中人！"
else
sql="select * from 用戶 where 姓名='" & name & "' and 門派='官府' and 身份='掌門'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess=name & "你不是掌門"
else
'		if a="a" then
select case a
case "a"
sql="update 用戶 set grade=6, 門派='官府' where 姓名='" & name2 & "'"
conn.execute sql
mess="你成功地把" & name2 & "招聘為官府的工作人員"
case "b"
sql="update 用戶 set 身份='弟子', 門派='遊俠' where id=" & id
conn.execute sql
mess="你成功地把" & id & "從官府開除了！"
case "c"
sf=request("sf")
sql="update 用戶 set grade='" & sf & "' where id=" & id
conn.execute sql
mess="你成功地把" & id & "的身份升（降）級為"&sf&"了！"
end select
end if
end if
conn.close
set conn=nothing
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>靈劍江湖</title></head>

<body background="../images/8.jpg">
<table border="1" align="center" width="350" cellpadding="10"
cellspacing="13" background="../images/12.jpg">
  <tr>
    <td height="89" background="../images/7.jpg"> 
      <table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="admin.asp">返回</a></p>
</td>
</tr>
</table>
    </td>
</tr>
</table>
<div align="center"> <br>
  <font color="#FFFFFF" size="2">★靈劍江湖版權所有★ </font></div>

</body>