<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
inthechat=Session("ljjh_inthechat")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
'if inthechat<>"1"  then Response.Redirect "manerr.asp?id=211"
lockname=Trim(Request.QueryString("id"))
lockip=Trim(Request.QuerySTring("ip"))
if lockname<>"" and lockip="" then Response.Redirect "manerr.asp?id=212"
if CStr(lockname)=CStr(nickname) then Response.Redirect "manerr.asp?id=213"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT ip,lockdate,locker FROM iplocktemp"
rs.open sql,conn,1,1
totalrec=rs.RecordCount
if totalrec>0 then
Dim show()
i=1
Do while Not rs.Eof
Redim Preserve show(i),show(i+1),show(i+2)
show(i)=rs("ip")
show(i+1)=rs("lockdate")
show(i+2)=rs("locker")
i=i+3
rs.MoveNext
Loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
t=s & ":" & f & ":" & m
sj=n & "-" & y & "-" & r & " " & t%><html>
<head>
<title>ＩＰ管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<Script Language=JavaScript>function ulk(ip){document.forms[1].unlockip.value=ip;document.forms[1].unlockwhy.focus();}</Script>
</head>
<body oncontextmenu=self.event.returnValue=false class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">【ＩＰ管理】</font></h1>
</div>
<hr noshade size="1" color=009900>
<b> 注意事項 </b> <br>
封鎖、解鎖 IP 的操作會被記錄在“聊務公開”欄中，供廣大聊友監督！<br>
被封鎖的 IP 在 <font color=red>50000</font> 分鐘內不能登錄。之後系統會自動解鎖該 IP。
<hr noshade size="1" color=009900>
<table border="1" cellpadding="5" cellspacing="0" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form method="post" action="manlockok.asp">
<tr>
<td>
<table border="0">
<tr>
<td>封鎖ＩＰ：</td>
<td>
<input type="text" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" name="lockip" value="<%=lockip%>" size="15" maxlength="15"<%if lockname<>"" then Response.Write " readonly"%>>
<%if lockname<>"" then Response.Write " <font color=red>(" & lockname & ")</font>"%>
<input type="hidden" name="lockname" value="<%=lockname%>">
</td>
</tr>
<tr>
<td>封鎖原因：</td>
<td>(&lt;=60字符)</td>
</tr>
<tr>
<td>
<select name="select" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" onChange="javascript:document.forms[0].lockwhy.value=this.value;document.forms[0].lockwhy.focus();">
<option value="" selected>自填</option>
<option value="不受歡迎人。">不歡迎</option>
<option value="所取的名字十分不雅。">不雅ID</option>
<option value="亂刷屏，警告又不聽。">亂刷屏</option>
<option value="在聊天室散布有悖倫理道德的言論。">不道德</option>
<option value="不遵守聊天規則，進行人身攻擊。">罵人</option>
<option value="在聊天室散布違反國家法律法規的言論。">違法</option>
</select>
</td>
<td>
<input type="text" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" name="lockwhy" maxlength="60" size="50">
</td>
</tr>
</table>
<div align="center">
<input type="submit" name="Submit" value="封鎖">
<input type="reset" name="reset" value="重寫">
<input type="button" value="放棄" onclick="javascript:history.go(-1)">
</div>
</td>
</tr>
</form>
</table>
<hr noshade size="1" color=009900>
<table border="1" cellpadding="5" cellspacing="0" bgcolor="E0E0E0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center">
<form method="post" action="manunlockok.asp">
<tr>
<td>
<table border="0">
<tr>
<td>解鎖ＩＰ：</td>
<td>
<input type="text" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" name="unlockip" size="15" maxlength="15" readonly>
</td>
</tr>
<tr>
<td>解鎖原因：</td>
<td>(&lt;=60字符)</td>
</tr>
<tr>
<td>
<select name="select" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" onChange="javascript:document.forms[1].unlockwhy.value=this.value;document.forms[1].unlockwhy.focus();">
<option value="" selected>自填</option>
<option value="封鎖時不小心搞錯了。">誤操作</option>
<option value="對方已經改過，請求解鎖IP。">悔過</option>
<option value="這個IP是網吧的，封鎖後都上不來了。">網吧</option>
<option value="這個IP是代理地址，封鎖後都上不來了。">代理</option>
</select>
</td>
<td>
<input type="text" name="unlockwhy" maxlength="60" size="50" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
</td>
</tr>
</table>
<div align="center">
<input type="submit" name="Submit" value="解鎖">
<input type="reset" name="reset" value="重寫">
<input type="button" value="放棄" onclick="javascript:history.go(-1)">
</div>
</td>
</tr>
</form>
</table>
<hr noshade size="1" color=009900>
<div align=center>共 <font color=red><%=totalrec%></font> 個 <a href=javascript:history.go(0)>刷新</a></div>
<table border="1" cellpadding="4" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="E0E0E0" align="center">
<tr bgcolor="#0099FF">
<td><font color="#FFFFFF">序號</font></td>
<td><font color="#FFFFFF">封 鎖 者</font></td>
<td><font color="#FFFFFF">封 鎖 時 間</font></td>
<td><font color="#FFFFFF">被封鎖的IP地址</font></td>
<td><font color="#FFFFFF">已封鎖時長</font></td>
<td><font color="#FFFFFF">立即解鎖</font></td>
</tr><%if totalrec>0 then
for i=1 to UBound(show) step 3%><tr>
<td><%=(i+2)/3%></td>
<td><%=show(i+2)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i)%></td>
<td><%=DateDiff("n",show(i+1),sj)&" "%>分鐘</td>
<td><a href="javascript:ulk('<%=show(i)%>')">Ｕ</a></td>
</tr><%next
end if%></table>
</body>
</html> 