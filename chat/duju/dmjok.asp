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
id=LCase(trim(request.querystring("id")))
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
	conn.close
	set rs=nothing
	set conn=nothing
  response.write "<script Language='Javascript'>alert('你不可以選擇大家或自已作為對手!');location.href = 'javascript:history.go(-1)';</script>"
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
duidu="【拒絕】["&info(0)&"]拒絕["& fromname &"]的搓麻將請求!"
duidu=duidu & "<br>.........搓麻將難搞得緊，我還要研究一下...."
else
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
db="DMJ.mdb"
'connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath(""&db&"")
'如果你的服務器支持jet4.0，請使用下載的連接方法，提高程序性能
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(""&db&"")
conn.open connstr 
sql="select * from dmj where ufrom='"& info(0) & "' or ufrom='"& fromname & "'"
Set Rs=conn.Execute(sql)
if not (rs.eof and rs.bof) then
  unpkname=rs("ufrom")
  rs.close
  conn.close
  response.write "<script Language='Javascript'>alert('" & unpkname & "還有牌局沒有結束，不能再次開局');location.href = 'javascript:history.go(-1)';</script>"
  abc=1
  duidu="【出錯】["&info(0)&"]不能接授["& fromname &"]的對賭請求!,因為["&unpkname&"]還有其他的牌局沒有結束。"
  msg="您還有牌局沒有結束，不能再次開局！"
else
  sql="select duz from dmj where uto='"& info(0) & "$下注' and ufrom='"& fromname & "$下注' order by id Desc"
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
  sql="delete * from dmj where uto='"& info(0) & "$下注' or ufrom='"& fromname & "$下注' or ufrom='"& info(0) & "$下注' or uto='"& fromname & "$下注'"
  Set Rs=conn.Execute(sql)

dim Allmjp
Allmjp="1萬|2萬|3萬|4萬|5萬|6萬|7萬|8萬|9萬|1筒|2筒|3筒|4筒|5筒|6筒|7筒|8筒|9筒|1索|2索|3索|4索|5索|6索|7索|8索|9索|東風|南風|西風|北風|紅中|白板|發財|"
Allmjp=Allmjp & Allmjp & Allmjp & Allmjp
mjpArr=split(Allmjp,"|")
'4*42=168

Randomize
for i= 1 to 14
	mjpArr=split(Allmjp,"|")
	Mjx= Int(ubound(mjpArr)* Rnd)
	strMjx=mjpArr(Mjx)
	Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
	myMj=myMj & strMjx & "|"
next

Randomize
for i= 1 to 13
	mjpArr=split(Allmjp,"|")
	Mjx= Int(ubound(mjpArr) * Rnd)
	strMjx=mjpArr(Mjx)
	Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
	youMj=youMj & strMjx & "|"
next

Allmjp_Akk=""
Randomize
for i= 1 to 109
	mjpArr = Split(Allmjp, "|")
	Mjx= Int(UBound(mjpArr) * Rnd)
	strMjx=mjpArr(Mjx)
	Allmjp=replace(Allmjp,strMjx & "|","",1,1,1)
	Allmjp_Akk=Allmjp_Akk & strMjx & "|"
next

sql="delete * from mjInfo where 莊家='" & info(0) & "' or 莊家='" & fromname & "' or 對手='" & info(0) & "' or 對手='" & fromname & "'"
Set Rs=conn.Execute(sql)
sql="insert into mjInfo(莊家,對手,麻將) values ('"& info(0) & "','"& fromname & "','"& Allmjp_Akk & "')"
Set Rs=conn.Execute(sql)
sql="select id from mjInfo where 莊家='"& info(0) & "'"
Set Rs=conn.Execute(sql)
mjID=Rs("id")
sql="insert into dmj(ufrom,uto,Mymj,duz,isMy,isGet,isFp,mjID) values ('"& info(0) & "','"& fromname & "','"& Mymj & "'," & xiazhu & ",true,false,true," & mjID & ")"
Set Rs=conn.Execute(sql)
sql="insert into dmj(ufrom,uto,Mymj,duz,isMy,isGet,isFp,mjID) values ('"& fromname & "','"& info(0) & "','"& youMj & "'," & xiazhu & ",false,false,false," & mjID & ")"
Set Rs=conn.Execute(sql)
duidu="<font color=green>["&info(0)&"]</font>同意搓麻將,<font color=green>[靈劍博士]</font>洗牌......唏哩嘩啦一陣暴響<br>"
duidu=duidu & "<font color=green>["&info(0)&"]</font>摸了14張牌<img src='duju/dpk/1.gif'>"
duidu=duidu & "<input type=button value='洗牌' onclick=javascript:window.open('duju/dmj-xp.asp?name="&info(0)&"','f3','width=380,height=210')  style=background-color:#86A231;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=dmjname><img src='duju/dpk/2.gif'>"
duidu=duidu & " <font color=green>["
duidu=duidu & fromname&"]</font>摸了13張牌<img src='duju/dpk/1.gif'><input type=button value='洗牌' onclick=javascript:window.open('duju/dmj-xp.asp?name="
duidu=duidu & fromname &"','f3','width=380,height=210') style=background-color:#86A231;color:FFFFFF;border: 1 double onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' name=dmjname><img src='duju/dpk/2.gif'>"
duidu=duidu & "<br>.....現在[" & info(0) & "]是莊家，請先打牌。<br>"
abc=1 
msg="你同意跟對方搓麻將了，請稍候！[靈劍總站長江湖總站]正在幫你們洗牌！"
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
sd(199)="<font color=#FFC01F>【搓麻將】</font>:"& duidu &"</font>"
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
<title>靈劍對賭</title>
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
function strMjp(cmj)
dim mj2
dim mjr
dim mjL
mjr=right(cmj,1)
if mjr=0 or cmj>40 then
mj2=replace(cmj,"10","東風")
mj2=replace(cmj,"20","南風")
mj2=replace(cmj,"30","西風")
mj2=replace(cmj,"40","北風")
mj2=replace(mj2,"41","紅中")
mj2=replace(mj2,"42","白板")
mj2=replace(mj2,"43","發財")
else
mjL=Left(cmj,1)
mjL=replace(mjL,"1","索")
mj2=replace(mjL,"2","筒")
mjL=replace(mjL,"3","萬")
mj2=mjr & mjL
end if
strMjp=mj2
end function
function intMjp(cmj)
dim mj2
dim mjL
mj2=cmj
mjL left(cmj,1)
if isNumeric(mjL) then mj2=right(cmj,1) & mjL
mj2=replace(mj2,"索","1")
mj2=replace(mj2,"筒","2")
mj2=replace(mj2,"萬","3")
mj2=replace(mj2,"風","0")
mj2=replace(mj2,"東","1")
mj2=replace(mj2,"南","2")
mj2=replace(mj2,"西","3")
mj2=replace(mj2,"北","4")
mj2=replace(mj2,"紅中","41")
mj2=replace(mj2,"白板","42")
mj2=replace(mj2,"發財","43")
intMjp=int(mj2)
end function
if to1="大家" or to1=fromname then
%>
<script language="JavaScript">window.close();</script>
<%end if%>


