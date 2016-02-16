
<body bgcolor="#660000">
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(3)<30 then Response.Redirect "../error.asp?id=460"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 用戶 where 姓名='"&info(0)&"' and 門派='遊俠'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('遊俠門的人不可以進行此操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select 掌門 from 門派 where 掌門='"&info(0)&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你已經是掌門了,還挑戰誰呢！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select 掌門 from 門派 where 門派='"&info(5)&"'",conn
zhangmen=rs("掌門")
rs.close
rs.open "select 等級,武功,魅力,知質 from 用戶 where 姓名='"&zhangmen&"'",conn
dj=rs("等級")
wg=rs("武功")
ml=rs("魅力")
zz=rs("知質")
rs.close
rs.open "select 身份,等級,武功,魅力,知質 from 用戶 where id="&info(9),conn
if rs("等級")<dj or rs("武功")<wg or rs("魅力")<ml or rs("知質")<zz then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的等級、武功、魅力、知質都能超越你派的掌門嗎？回去再好好練練吧！');window.close()}</script>"
	Response.End 
end if
conn.execute("update 用戶 set 身份='弟子',grade=1 where 姓名='"&zhangmen&"'")
conn.execute("update 用戶 set 身份='掌門',grade=5 where id="&info(9))
conn.execute("update 門派 set 掌門='"&info(0)&"' where 門派='"&info(5)&"'")
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
sd(195)=info(0)
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="對"
sd(199)="<font color=FFD7F4>【挑戰掌門】"& info(0) &"</font><font color=FFD7F4>以優秀的武功超越了本派掌門成為 "&info(5)&" 的新掌門！</font></font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{alert('你以優秀的武功成為本派新任掌門,請退出江湖重新進去！');window.close();}</script>"
Response.End
%>