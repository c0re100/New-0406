<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if   info(2)>=7  then
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name1=request("name")
fg=request("fg")
my=request("my")
rs.open "select 姓名1,理由 from 離婚 where 姓名1='" & name1 & "'and 姓名2='" & my &"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你所要判決離婚的人["&my&"]不存在！！');history.go(-1);</script>"
	response.end
end if
liu=rs("理由")
'rs.close
'rs.open "select * from 用戶 where 姓名='" & name1 & "'",conn
'if rs.eof or rs.bof then
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script Language=Javascript>alert('提示：你所要判決離婚的人["&name1&"]不存在！！');history.go(-1);</script>"
'	response.end
'end if
conn.execute "update 用戶 set 配偶='無',銀兩=銀兩+10000,結婚時間=date(),小孩='無' where 姓名='" & my & "'"
conn.execute "update 用戶 set 配偶='無',結婚時間=date(),小孩='無' where 姓名='" & name1 & "'"
conn.execute "delete * from 離婚 where 姓名1='" & name1 & "' or 姓名2='" & name1 & "' "
message="[" & name1 & "]和[" & my & "]經官府判決離婚成功，理由是<font color=FFFDAF>{"&liu&"}</font>,為他們重獲自由鼓掌！"
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
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
sd(199)="<font color=red>【離婚消息】"& message &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock%>
<body bgcolor="#000000">
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