<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if
to1=request.form("to1")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(to1)&" ")=0 then
Response.Write "<script Language=Javascript>alert('所要施法的人不在聊天室！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if
if info(3)<10 then
Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上才可以使用此法術！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,狀態,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<6 then
rs.close
set rs=nothing
conn.close
set conn=nothing
ss=6-sj
Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
if rs("狀態")="死" then
Response.Write "<script language=JavaScript>{alert('你已經在假死狀態，當心掉線，掉線了就得在首頁復活了！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
if rs("法力")<1000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得1000點啊！');location.href = 'javascript:history.go(-1)';}</script>"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.End
end if
rs.close
rs.open "select grade FROM 用戶 WHERE 姓名='"&to1&"'",conn
conn.execute "update 用戶 set 狀態='死',法力=法力-1000,操作時間=now() where id="&info(9)
jiasi="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>突然跑到<font color=red>"&to1&"</font></a>家門口,一不小心被石頭拌了一下摔在地上一動也不動，看樣子可能是運氣不好一下就死翹翹了吧!"

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
sd(194)=to1
sd(195)=info(0)
sd(199)="<font color=CFE0B0>【假死術】"& jiasi &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>location.href = 'javascript:history.go(-1)';</script>"
Response.End
%>