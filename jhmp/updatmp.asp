<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
comment=request.form("comment")
shihe=trim(request.form("shihe"))
zhangmen=trim(request.form("zhangmen"))
s=Request.QueryString("subject")
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
if chuser(zhangmen) then Response.Redirect "../error.asp?id=120"
if shihe<>"男" and shihe<>"女" and shihe<>"所有" then Response.Redirect "../error.asp?id=442"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if zhangmen="無" then
rs.open "select * FROM 門派 WHERE 門派='"&s&"' and 掌門='"&info(0)&"' ",conn
if rs.EOF or rs.BOF then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('抱歉你不是掌門不能操作！');history.back();</script>"
Response.End
else
conn.execute "Update 門派 Set 適合='" & shihe & " ',簡介='" & comment & " ' Where 門派='" & s & "'"
end if
else
rs.open "select id FROM 用戶 WHERE 門派='"&s&"' and 姓名='"&zhangmen&"' ",conn
if rs.EOF or rs.BOF then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('要傳位的人不是你派弟子！');history.back();</script>"
Response.End
else
conn.execute "Update 門派 Set 掌門='"&zhangmen&"' Where 門派='" & s & "'"
conn.execute "Update 用戶 Set 門派='" & s & "',身份='掌門',grade=5 Where 姓名='"&zhangmen&"'"
conn.execute "Update 用戶 Set 門派='" & s & "',身份='弟子',grade=1 Where 姓名='"&info(0)&"'"
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
sd(199)="<font color=FFD7F4>【系統】[" & s & "]的掌門:"& info(0) &"</font><font color=FFD7F4>把掌門寶座傳給了 "&zhangmen&" ！</font></font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end if
end if
Response.Redirect "../ok.asp?id=701"
%> 