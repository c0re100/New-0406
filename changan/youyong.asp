<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('你不能進行操作，進行此操作必須進入聊天室！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
my=info(0)%>
<!--#include file="data1.asp"-->
<%sql="select * from 房屋 where 戶主='" & my & "' or 伴侶='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then%>
<script language=vbscript>
MsgBox "您還沒有買房子呢。"
location.href = "fangwu.asp"
</script>
<%else
if day(rs("遊泳"))<day(now()) and month(rs("遊泳"))<month(now()) and year(rs("遊泳"))<year(now()) and isnull(rs("遊泳")) then
sql="update 房屋 set 遊泳=now() where 戶主='"& my &"' or 伴侶='" & my & "'"
set rs=conn.execute(sql)
set rs=nothing	
	conn.close
	set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 用戶 where id="&info(9),conn
sql="update 用戶 set 魅力=魅力+100,體力=體力+10000 where id="&info(9)
set rs=conn.execute(sql)
Response.Redirect "../ok.asp?id=602"
else
%><script language=vbscript>
MsgBox "您今天遊過泳了啊。"
location.href = "xiaowu7.asp"
</script>
<%end if
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>