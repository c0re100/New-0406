<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nowinroom=session("nowinroom")
if Instr(LCase(Application("ljjh_useronlinename"&nowinroom)),LCase(" "&info(0)&" "))=0 then 
	Session("ljjh_inthechat")="0" 
	Response.Redirect "../chaterr.asp?id=001" 
end if 

chatroomsn=session("nowinroom")

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
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
id=trim(request.querystring("id"))
fromname=trim(request.querystring("from1"))
to1=trim(request.querystring("to1"))
yn=trim(request.querystring("yn"))
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select 姓名 FROM 用戶 WHERE id="&fromname,conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('對不起，對方的名字不存在。');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
fromname=rs("姓名")
rs.close
if to1="大家" or to1=fromname then 
response.write "<script Language='Javascript'>alert('您現在沒有參與牌局，請向你的對手使用\""/打撲克 1000 \""命令要求開局。');location.href = 'javascript:history.go(-1)';</script>"
response.end
end if
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
if info(0)<>to1 then
  msg="人家沒有邀請你呀！"
  abc=0
elseif yn=0 then
  msg="你拒絕人家的對賭了！"
  abc=1
  duidu="【拒絕】["&info(0)&"]拒絕["& fromname &"]的打撲克牌請求!"
  duidu=duidu & "<br>.........打撲克難搞得緊，我還要研究一下..."
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="dpk.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dpk where ufrom='"& info(0) & "' or ufrom='"& fromname & "'"
Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
unpkname=rs("ufrom")
rs.close
conn.close
set rs=nothing
set conn=nothing
response.write "<script Language='Javascript'>alert('" & unpkname & "還有牌局沒有結束，不能再次開局');location.href = 'javascript:history.go(-1)';</script>"
abc=1
duidu="【出錯】["&info(0)&"]不能接授["& fromname &"]的對賭請求!,因為["&unpkname&"]還有其他的牌局沒有結束。"
msg=unpkname & "還有牌局沒有結束，不能再次開局！"
else

sql="select duz from dpk where uto='"& info(0) & "$下注' and ufrom='"& fromname & "$下注' order by id Desc"
Set Rs=conn.Execute(sql)
	If rs.eof and rs.bof Then
	rs.close
	conn.close
	set rs=nothing
	set conn=nothing
response.write "<script Language='Javascript'>alert('錯誤:沒有這個賭局!');location.href = 'javascript:history.go(-1)';</script>"
	response.end
	end If

xiazhu=Rs("duz")
sql="delete * from dpk where uto='"& info(0) & "$下注' or ufrom='"& fromname & "$下注' or ufrom='"& info(0) & "$下注' or uto='"& fromname & "$下注'"
Set Rs=conn.Execute(sql)

for i= 1 to 18
Randomize
apk = Int(14 * Rnd+1)
Randomize
bbb= Int(4 * Rnd+1)
if bbb=1 then bbb2="紅"
if bbb=2 then bbb2="黑"
if bbb=3 then bbb2="梅"
if bbb=4 then bbb2="方"
dpk=bbb2 & apk
dpk=convpk(dpk)
if instr(frpk,dpk & "|")=0 then
frpk=frpk & dpk & "|"
else 
i=i-1
end if
next
for i= 1 to 18
Randomize
apk = Int(14 * Rnd+1)
Randomize
bbb= Int(4 * Rnd+1)
if bbb=1 then bbb2="紅"
if bbb=2 then bbb2="黑"
if bbb=3 then bbb2="梅"
if bbb=4 then bbb2="方"
dpk=bbb2 & apk
dpk=convpk(dpk)
if instr(frpk,dpk & "|")=0 and instr(topk,dpk & "|")=0 then
topk=topk & dpk & "|"
else 
i=i-1
end if
next
sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& info(0) & "','"& fromname & "','"& topk & "',false," & xiazhu & ",90)"
Set Rs=conn.Execute(sql)
sql="insert into dpk(ufrom,uto,pk,fp,duz,oldpn) values ('"& fromname & "','"& info(0) & "','"& frpk & "',true," & xiazhu & ",90)"
Set Rs=conn.Execute(sql)
duidu="<font color=green>["&info(0)&"]</font>同意打撲克,<font color=green>[靈劍博士]</font>開始分牌......<br>"
duidu=duidu & "<font color=green>["&info(0)&"]</font>抓了18張牌<img src='duju/dpk/1.GIF'>"
duidu=duidu & "<input type=button value='洗牌' onclick=javascript:window.open('duju/dpk-xp.asp?name="&info(0)&"','f3','width=380,height=210') style='background-color:#86A231;color:FFFFFF;border: 1 double' onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=dpkname><img src='duju/dpk/2.GIF'>"
duidu=duidu & " <font color=green>["
duidu=duidu & fromname&"]</font>抓了18張牌<img src='duju/dpk/2.GIF'><input type=button value='洗牌' onclick=javascript:window.open('duju/dpk-xp.asp?name="
duidu=duidu & fromname &"','f3','width=380,height=210')  style=background-color:#86A231;color:FFFFFF;border: 1 double  onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=dpkname><img src='duju/dpk/1.GIF'>"
duidu=duidu & "<br>.....根據牌局規定，請[" & info(0) & "]先發牌。<br>"
abc=1 
msg="你同意跟對方打撲克了，請稍候！[江湖博士]正在幫你們洗牌！"
set conn=nothing
set rs=nothing
end if
end if
if abc=1 then
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)=info(0)
sd(196)="FFC01F"
sd(197)="FFC01F"
sd(198)="對"
sd(199)="<font color=#FFC01F>【打撲克】</font>:"& duidu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock



end if
%>
<html>
<head>
<style>
a:link {text-decoration:none; font-size:14;color:#000000}
a:hover {text-decoration:underline;font-size:14; color:#000000; background:#ffffff}
a:visited {text-decoration:none;font-size:14; color:#000000}
td {font-size:9pt; color:#ff0000; line-height:16pt}
</style>
<META http-equiv='content-type' content='text/html; charset=big5'>
<title>對賭</title>
</head>
<body bgcolor=#0099FF ><table>
<tr>
<td> <font color="red"> </font><br>
<table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" bgcolor="#FFFFFF">
<tr> 
<td height="115"> 
<table border="0" bgcolor="#009900" cellspacing="0" cellpadding="2" width="361">
<tr> 
<td width="324" height="9"><font color="FFFFFF" face="Wingdings">z</font><font color="FFFFFF">靈劍提示</font><font color="red" size=2> 
</font></td>
<td width="29" height="9"> 
<table border="1" bordercolorlight="666666" bordercolordark="FFFFFF" cellpadding="0" bgcolor="E0E0E0" cellspacing="0" width="18">
<tr> 
<td width="16"><b><a href="<%if id="200" then%>javascript:window.close();<%else%>javascript:history.go(-1)<%end if%>" onMouseOver="window.status='';return true" onMouseOut="window.status='';return true" title="返回"><font color="000000">×</font></a></b></td>
</tr>
</table>
</td>
</tr>
</table>
<table border="0" width="359" cellpadding="4">
<tr> 
<td width="59" align="center" valign="top" height="29"><font face="Wingdings" color="#FF0000" style="font-size:32pt">J</font></td>
<td width="278" height="29"> <font color="red" size=2> 
<%=msg%>
</font></td>
</tr>
<tr> 
<td colspan="2" align="center" valign="top" height="58"> 
<input type=button value='關閉' onClick="javascript:window.close()" style="background-color:3366FF;color:FFFFFF;border: 1 double" onMouseOver="this.style.color='FFFF00'" onMouseOut="this.style.color='FFFFFF'" name="button223">
</td>
</tr>
</table>
</td>
</tr>
</table>
</body>
<%
function convpk(dpk)
if dpk="紅14" then dpk="大王"
if dpk="方14" then dpk="大王"
if dpk="梅14" then dpk="小王"
if dpk="黑14" then dpk="小王"
dpk=replace(dpk,"13","K")
dpk=replace(dpk,"12","Q")
dpk=replace(dpk,"11","J")
dpk=replace(dpk,"10","十")
dpk=replace(dpk,"1","A")
dpk=replace(dpk,"十","10")
convpk=dpk
end function 
if to1=jiutian_dabusi then
%>
<script language="JavaScript">window.close();</script>
<%end if%>
