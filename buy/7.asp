<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "./error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
my=info(0)
rs.open "select 操作時間,會員等級,銀兩 from 用戶 where id=" & info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<1 then
	s=3-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對不起系統限制，請等["&s&"秒鐘]再操作！');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if	
if rs.eof or rs.bof then
	response.write "喵"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
else
tl=rs("銀兩")
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<center>
<%
if tl<1 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "錯誤！你帶不夠錢來喔！"
location.href = "javascript:history.back()"
</script>
<%
else
conn.execute "update 用戶 set 銀兩=銀兩-10000000,攻擊=攻擊+1000000,操作時間=now() where id="&info(9)
	rs.close

set rs=nothing
conn.close
set conn=nothing
message="" & my & "購買成功,重新登入！"
		
end if
end if%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>江湖總站:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br><center>
<p></p>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='index.asp'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>