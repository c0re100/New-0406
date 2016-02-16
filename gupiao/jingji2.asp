<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="fun.inc"-->
<%Response.Expires=0
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "請進入聊天室再進入股票！"
window.close()
</script>
<%Response.End
end if
name=trim(request.form("name"))
jjr=request.form("jjr")
%>
<!--#include file="jhb.asp"-->
<%
sql="select * from 客戶 where 帳號='"&name&"'"
set rs=conn.execute(sql)
if rs.eof then
	sql="insert into 客戶(帳號,經紀人,資金)values('" & name & "','"&jjr&"',0)"
	conn.execute sql
	mess="股票帳戶申請成功！我們安排了金牌經紀人"&jjr&"為您服務."
else
	sql="update 客戶 set 經紀人='"&jjr&"' where 帳號='"&name&"'"
	conn.execute sql
	mess="我們安排了新的經紀人"&jjr&"為您服務."
end if
conn.close
%>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=mess%></p>
<p align=center><a href=jingji.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>