<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="jhb.asp"--><%
jin=request.form("money")
jin=abs(int(jin))
mess=""
Jname=info(0)
sql="select * from 客戶 where 帳號='" & Jname & "' "
set rs=conn.execute(sql)
yin=rs("資金")
if jin>yin then
mess="你的股票經紀人大聲說，我沒有幫你賺這麼多錢呀？"
else
sql="update 客戶 set 資金=資金-"& jin & " where 帳號='" & Jname & "'"
conn.execute sql
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩 from 用戶 where id="&info(9),conn
money=rs("銀兩")
moneyok=int(money)+abs(fn1)
if moneyok>2000000000 then
	Response.Write "<script language=JavaScript>{alert('你的銀兩就快超過了20億了，為防止丟錢現像請您少取些吧！');location.href = 'javascript:history.go(-1)';}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sql="update 用戶 set 銀兩=銀兩+1 where 姓名='" & Jname & "'"
conn.execute sql
mess="你從股票帳戶取出了1兩，拿好了啊"
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>

<head>
<title></title>
</head>

<body text="#000000" background="../jhimg/bk_hc3w.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=qu.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>