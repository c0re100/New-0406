<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
grade=Int(info(2))
nickname=info(0)
id=Trim(Request.QueryString("id"))
id2=LCase(id)
if nickname="" then Response.Redirect "../error.asp?id=420"
mypai=info(5)
if grade<7 then Response.Redirect "../error.asp?id=482"
if id2="" then Response.Redirect "../error.asp?id=483"
total=Application("ljjh_chatrs")
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
ubo=UBound(onlinelist)
dim show()
js=1
for i=1 to ubo step 6
if Instr(LCase(onlinelist(i+1)),id2)<>0 or Instr(onlinelist(i+2),id2)<>0 then
Redim Preserve show(js),show(js+1),show(js+2),show(js+3)
show(js)=onlinelist(i+1)
show(js+1)=onlinelist(i+2)
show(js+2)=onlinelist(i+4)
show(js+3)=onlinelist(i+5)
js=js+4
end if
next
set onlinelist=nothing
totalrec=Int((js-1)/4)
p=1
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m%><html>
<head>
<title>聊友信息查找</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
<script language=javascript>window.moveTo(100,50);window.resizeTo(screen.availWidth*2/3,screen.availHeight*3/4);</script>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#FFFFFF" class=p150>
<div align="center">
<h1><font color="#000000" face="隸書" size="6">炸 人 操 作</font></h1>
<a href="javascript:history.go(0)">刷新</a> </div>
<hr noshade size="1" color=009900 width="70%">
<div align="center"><b> 在線人員 </b> 共 <font color=red><%=total%></font> 人，找到包含 <font color=red><%=id%></font>
的用戶名 <font color=red><%=totalrec%></font> 個。 </div>
<hr noshade size="1" color=009900 width="70%">
<div align=center>本頁刷新時間：<font color="#FF0000"><%=sj%></font><br></div>
<table border="1" cellspacing="0" bordercolorlight="#999999" bordercolordark="#FFFFFF" cellpadding="4" bgcolor="E0E0E0" align="center" width="416">
<tr bgcolor="#000000" align="center">
<td><font color="#FFFFFF">序號</font></td>
<td><font color="#FFFFFF">用 戶 昵 稱</font></td>
<td><font color="#FFFFFF">性別</font></td>
<td><font color="#FFFFFF">進 入 時 間</font></td>
<td><font color="#FFFFFF">未存點</font></td>
<td><font color="#FFFFFF">炸彈</font></td>
</tr>
<%if grade<3 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=3 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=4 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=5 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade=6 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%> <%next
end if
end if
if grade>6 then
if totalrec>0 then
j=0
for i=1 to UBound(show) step 4%>
<tr align="center">
<td><%Response.Write p+j
j=j+1%></td>
<td><%=show(i)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i+2)%></td>
<td><%=DateDiff("n",show(i+3),sj)%> 分</td>
<td><a href="manbomb.asp?id=<%=server.URLEncode(show(i))%>&ip=<%=show(i+1)%>">轟炸</a></td>
</tr>
<%next
end if
end if%>
</table>
<div align="center"><br>
注意：此功能隻用於那些經常搗亂卻不悔改的聊友!<br>
此記錄將記載於數據庫中，任何網管瞎炸人的都將會受到處罰 </div>
</body>
</html> 