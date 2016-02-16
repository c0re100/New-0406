<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(5)="官府" and info(2)>=7  then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name1=request("name")
fg=request("fg")
my=request("my")
rs.open "select 姓名1 from 離婚 where 姓名1='" & name1 & "'and 姓名2='" & my &"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你所要判決離婚的人["&my&"]不存在！！');history.go(-1);</script>"
	response.end
end if
rs.close
rs.open "select ID from 用戶 where 姓名='" & name1 & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你所要判決離婚的人["&name1&"]不存在！！');history.go(-1);</script>"
	response.end
end if
conn.execute "delete * from 離婚 where 姓名1='" & name1 & "' or 姓名2='" & name1 & "' "
message="" & my & "與" & name1 & "經官府判決離婚無效，解除離婚登記！"
conn.close
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
sd(195)="大家"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="對"
sd(199)="<font color=FFFDAF>【消息】'"& message &"'</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock%>
<body bgcolor="#000000" background="../jhimg/bk_hc3w.gif">
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table width=100%>
<tr><td>
<p align=center style='font-size:14;color:red'><%=message%></p>
<p align=center><a href=Yuanou.asp>返回</a></p>
</td></tr>
</table>
</td></tr>
</table>
</body>
<%end if%> 