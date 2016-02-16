<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
my=info(0)
rs.open "select 操作時間,體力,內力 from 用戶 where id="& info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對不起系統限制，請等["&s&"秒鐘]再操作！');location.href = 'javascript:history.go(-1)'}</script>"
	Response.End
end if	
if rs.eof or rs.bof then
	response.write "你不是江湖中人，不能進入江湖武館"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
else
tl=rs("體力")
nl=rs("內力")
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
if tl<550 or nl<260 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "錯誤！你目前的體力或內力值不夠，回去多加練練馬步蹲襠再來吧！"
location.href = "javascript:history.back()"
</script>
<%
else
conn.execute "update 用戶 set 體力=體力-550,內力=內力-260,武功=武功+100,防御=防御+10,操作時間=now() where id="&info(9)
rs.close
rs.open "select 等級,武功,武功加,防御,防御加 from 用戶 where id=" & info(9),conn
wgj=(rs("等級")*1280+3800)+rs("武功加")
fyj=(rs("等級")*1120+3000)+rs("防御加")
if rs("武功")>wgj then
conn.execute "update 用戶 set 武功=" & wgj & ",操作時間=now() where id="&info(9)
end if
if rs("防御")>fyj then
conn.execute "update 用戶 set 防御=" & fyj & ",操作時間=now() where id="&info(9)
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
message="" & my & "辛辛苦苦學了一天，你獲得武功25防御10！消耗體力550內力260！"
		
end if
end if%>
<table border=1 bgcolor="#948754" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#C6BD9B>
<table height="260">
<tr><td height="37">
<font color="#000000"><strong>江湖武館:</strong></font>
<tr>
<td height="182" valign="top"><%=message%><Br><Br><center>
</td>
</tr>
<td align=center height="29">&nbsp;
<div align="right">
<input type=button value="返 回" onclick="location.href='xuetang.htm'">
</div>
</td></tr>
</table>
</td></tr>
</table>
</body>
</html>