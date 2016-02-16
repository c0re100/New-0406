<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="fun.inc"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "請進入聊天室再進入股票！"
window.close()
</script>
<%Response.End
end if

name=info(0)
%>
<!--#include file="jhb.asp"-->
<%
	sql="select * from 客戶 where 帳號='" & name & "'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
	set rs=nothing
conn.close
set conn=nothing%>
<script language="vbscript">
  alert "你還沒有開戶呢！"
  location.href = "jingji.asp"
</script>
<%	   
    else
    jjr=rs("經紀人")
%>
<html>
<head>
<title>經紀人辦公室</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style>
td{color:#000000}
p{font-size:16;color:red}
</style>
</head>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>您的股票經濟人<%=jjr%>在此為您服務</p><br>
<a href=cun.asp>存錢進股票帳戶</a><br><a href=qu.asp>從股票帳戶提錢</a><br><a href=cha.asp>查看最近股票買賣情況</a><br><a href=huan.asp>換經紀人</a>  
<br>  
<p align=center><a href=jingji.asp>離開</a></p>
</table></table>  

		
</body>
</html>
<%
end if
end if
%>