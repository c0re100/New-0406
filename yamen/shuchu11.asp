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
'����Τ�
sql="SELECT * FROM �Τ� WHERE �m�W='" & my & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
mess=my&"�A�A����ާ@�I"
else
if rs("����")="�x��" and rs("����")>"6" then
UID=rs("ID")
sql="select * from �Τ� where ���A='�c' or ���A='��' and ����<>'�x��' and id=" & id
set rs=conn.execute(sql)
if not rs.eof and not rs.bof then
sql="update �Τ� set ���A='���`',�n��=now() where id=" & id
conn.execute sql
mess=my&"�A�A�w�g��Ӹo������F�I"
pos=application("chatpos")
chats=application("chats")
sj=formatdatetime(time(),3)
sj="<font color=#FF00FF size=1>(" & sj & ")</font>"
application.lock
chats(pos,1)=""
chats(pos,2)=red
chats(pos,3)=""
chats(pos,4)="�j�a"
chats(pos,5)=addsays
chats(pos,6)="�i����j"
chats(pos,7)="<font color=red><b>" & rs("�m�W") & "�Q�L�o����</b>(����xID="&UID&")</font>"
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
mess="�S���o�ӥǤH�άO���H�O�x�����H"
end if
else
mess=mye&"�A�A�L���v�O�I�I�I"
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
<p align="center"><a href="listlao.asp">��^</a></p>
</td>
</tr>
</table>
</td>
</tr>
</table>

</body> 