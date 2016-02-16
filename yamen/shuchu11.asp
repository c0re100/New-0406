<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<!--#include file="../config.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(0)
my=ljjh_nickname
id=request("id")
'校驗用戶
sql="SELECT * FROM 用戶 WHERE 姓名='" & my & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
mess=my&"，你不能操作！"
else
if rs("門派")="官府" and rs("身份")>"6" then
UID=rs("ID")
sql="select * from 用戶 where 狀態='牢' or 狀態='獄' and 門派<>'官府' and id=" & id
set rs=conn.execute(sql)
if not rs.eof and not rs.bof then
sql="update 用戶 set 狀態='正常',登錄=now() where id=" & id
conn.execute sql
mess=my&"，你已經把該罪犯釋放了！"
pos=application("chatpos")
chats=application("chats")
sj=formatdatetime(time(),3)
sj="<font color=#FF00FF size=1>(" & sj & ")</font>"
application.lock
chats(pos,1)=""
chats(pos,2)=red
chats(pos,3)=""
chats(pos,4)="大家"
chats(pos,5)=addsays
chats(pos,6)="【釋放】"
chats(pos,7)="<font color=red><b>" & rs("姓名") & "被無罪釋放</b>(執行官ID="&UID&")</font>"
chats(pos,8)=towhoway
chats(pos,0)=sj
if pos<MaxTalk then
application("chatpos")=pos+1
else
application("chatpos")=1
end if
application("chats")=chats
application.unlock
else
mess="沒有這個犯人或是此人是官府的人"
end if
else
mess=mye&"，你無此權力！！！"
end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000">

<table border="1" bgcolor="#BEE0FC" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#CCE8D6">
<table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="listlao.asp">返回</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body> 