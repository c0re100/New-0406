<%@ LANGUAGE = VBScript codepage ="950" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if session("TWT_ARR_ArgALL")="" then response.end
TWT_ArrArg=split(session("TWT_ARR_ArgALL"),"=")
nickname=TWT_ArrArg(0)
grade=TWT_ArrArg(2)
myid=TWT_ArrArg(1)
if grade>10 then
sql="delete * FROM 用戶 where id=" & myid
set Rs=conn.Execute(sql) 
conn.close 
session.Abandon 
Response.write "好好玩,去死吧!我們不歡迎妳呢個黑客。" 
response.end 
end if
%>
<HTML><HEAD><TITLE>天外天江湖--官府衙門</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="pic/bg.jpg">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="32%" height="31">
<tr align="center"> 
<td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">天外天衙門</font></b></td>
</tr>
</table>
<br>
<table border=0 cellpadding=0 cellspacing=0 width="601" align="center">
<tbody> 
<tr> 
<td width="11"><img src="pic/empty.gif" width="1" height="1"></td>
<td noWrap colspan="3">&nbsp; </td>
</tr>
<tr> 
<td width="11">　</td>
<td width="618">　</td>
<td width="36">　</td>
<td width="20">　</td>
</tr>
<tr> 
<td rowspan=2 width="11">　</td>
<td class=bg1 colspan=2 height="73" 
style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
<table width='545' align=left cellspacing=1 border=1 cellpadding=0 bgcolor='#FFFFFF' bordercolor="#000000">
<tr bgcolor="#ffffff"> 
<td align=center width="32" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">ID</font></div>
</td>
<td align=center width="65" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">姓 名</font></div>
</td>
<td align=center width="48" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">等 級</font></div>
</td>
<td align=center width="88" height="21" bgcolor="#FFFFFF"> <font size="2">支持率</font> </td>
<td align=center width="90" height="21" bgcolor="#FFFFFF"> <font size="2">投票</font> 
</td>
<td align=center width="241" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">站長操作</font></div>
</td>
</tr>
<% 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="SELECT * FROM 用戶 where 門派='六扇門' and 身份<>'掌門' order by grade DESC"
Set Rs=conn.Execute(sql)
do while not rs.eof and not rs.bof
piao=rs("piao")
%>
<tr bgcolor="#C4DEFF"> 
<td width="32" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"><%=rs("id")%></font></div>
</td>
<td width="65" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#ff0000"><%=rs("姓名")%></font></div>
</td>
<td width="48" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"><%=rs("grade")%></font></div>
</td>
<td width="88" bgcolor="#FFFFFF" height="2"> <font size="2" color="#000000"><img src=bar.gif width=<%=piao%> height=10><%=piao%>票</font> 
</td>
<td width="95" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"><a href=adddata.asp?id=<%=rs("id")%>>支持票</a>|<a href=adddata2.asp?id=<%=rs("id")%>>抗議票</a></font></div>
</td>
<td width="241" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"> 
<%if rs("grade")<=grade and rs("身份")<>"掌門" then%>
<a href=admin1.asp?id=<%=rs("id")%>&a=c&sf=<%=rs("grade")+1%>>升級</a> 
| <a href=admin1.asp?id=<%=rs("id")%>&a=c&sf=<%=rs("grade")-1%>>降級</a> 
| <a href=admin1.asp?id=<%=rs("id")%>&a=b>開除</a> | 
<%end if%>
<a href=admin2.asp?id=<%=rs("id")%> title="用於被打死後取回六扇門身份！">取回身份</a></font></div>
</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<p><font size="2"> <br>
<br>
</font></p>
<p>　</p>
<p>　</p>
</td>
<td valign=top width="20" height="73">　</td>
</tr>
<tr> 
<td class=bg1 colspan=2 
style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
<form method=POST action='admin1.asp?a=a'>
<p><font size="2" color="#FFFFFF">招收：</font><font size="2"> 
<input name=name2 type=text size=10>
<input value="確定" type=submit name="submit2">
</font> </p>
</form>
<p>
<font color="#FF0000">&nbsp;關於投票：投支持票，對方的支持率加一，投反對票，對方的支持率減一<br>
<font color="#FF0000">&nbsp;站長會定期檢查支持率為負數的管理員，並予以處罰<br><br>
<span class="unnamed1"><font size="2" color="#FFFFFF">六扇門身份分為6--10級，按照身份的高低擁有不同的權力！其中：<br>
<br>
警告與點死穴（踢人）為基本權力<br>
<br>
<font color="#FF0000">&nbsp;6級</font> 以上可以逮捕<br>
<font color="#FF0000">&nbsp;7級</font> 以上可以發錢<br>
<font color="#FF0000">&nbsp;8級</font> 以上可以抓人坐牢，24小時不可進江湖<br>
<font color="#FF0000">&nbsp;9級</font> 以上可清藥和對犯人進行斬首<br>
<font color="#FF0000">10級</font> 可對犯人進行斬首進管理目錄，可對六扇門中人升降級<br>
<font color="#FF0000">掌門</font> 可以開除六扇門中人</font></span><font color="#FFFFFF">, 
可以把別人升到十級。</font></p>
</td>
<td width="20">　</td>
</tr>
<tr> 
<td width="11">　</td>
<td colspan=2>　</td>
<td width="20">　</td>
</tr>
</tbody> 
</table>
<p>　</p>
</BODY></HTML>