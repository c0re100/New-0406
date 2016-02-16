<%@ codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="chkhk.asp"-->
<%
if session("TWT_ARR_ArgALL")="" then response.end
TWT_ArrArg=split(session("TWT_ARR_ArgALL"),"=")
nickname=TWT_ArrArg(0)
grade=TWT_ArrArg(2)
myid=TWT_ArrArg(1)
set TWT_ArrArg=nothing

userip=Request.ServerVariables("REMOTE_ADDR")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr

jingao=false
sql="SELECT value FROM system where name='toupaio'"
set rs=conn.execute(sql)
if rs.eof then
	sql="SELECT name,value FROM system"
	sql="insert into [system]([name],[value]) values ('toupaio','0')"
	conn.execute(sql)
else
	ipreip=rs(0)
	if ipreip=userip then
        jingao=true
	else
	sql="update [system] set [value]='"&userip&"' where [name]='toupaio'"
	conn.execute(sql)
	end if
end if
        rs.close

twtjh_tp=request.cookies("yescnet")("tp")
if tp<>"tp" then
	Response.cookies("yescnet")("tp")="tp"
	Response.cookies("yescnet").Expires=now+1
else
	jingao=true
end if

id=request("id")
sql="select piao,grade from 用戶 where ID=" & myid
set rs=conn.execute(sql)
if rs("piao")>=1 then
rs.close
conn.close
set conn=nothing
set rs=nothing
%>
<script language=vbscript>
MsgBox "您已經投了票，謝謝您的支持（每個玩家最多隻能投一票）！"
location.href = "admin.asp"
</script>
<%
elseif jingao then
rs.close
conn.close
set conn=nothing
set rs=nothing
call showchat("<font color=ff0000>【繫統報警】</font>" & nickname & "在官府衙門試圖以作弊方式給ID為" & id & "的管理員投反對票，被繫統攔截")
%>
<script language=vbscript>
MsgBox "不要這樣無聊好嗎？這次是警告，下次有你好看！"
location.href = "admin.asp"
</script>
<%
else
if rs("grade")<2 then
rs.close
conn.close
set conn=nothing
set rs=nothing
%>
<script language=vbscript>
MsgBox "您的等級不夠不能投票！"
location.href = "admin.asp"
</script>
<%else
sql="SELECT piao FROM 用戶 WHERE 門派='六扇門' and ID=" & id
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
conn.close
set conn=nothing
set rs=nothing
%>
<script language=vbscript>
MsgBox "您投誰啊？好像六扇門裡沒有這個人吧！"
location.href = "admin.asp"
</script>
<%else
sql="update 用戶 set piao=piao-1 where id=" & id
conn.execute sql
sql="update 用戶 set piao=piao+1 where id=" & myid
conn.execute sql
conn.close
set rs=nothing
set conn=nothing
%>
<script language=vbscript>
MsgBox "謝謝您，你已經成功的投了一個反對票，對方的支持率會減一" & chr(10) & "站長會下期檢查支持率為負數的管理員並予以處罰！"
location.href = "admin.asp"
</script>
<%
'end if
end if
end if
end if
%>
<!--#include file="chat.asp"-->