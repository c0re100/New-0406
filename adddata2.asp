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
sql="select piao,grade from �Τ� where ID=" & myid
set rs=conn.execute(sql)
if rs("piao")>=1 then
rs.close
conn.close
set conn=nothing
set rs=nothing
%>
<script language=vbscript>
MsgBox "�z�w�g��F���A���±z������]�C�Ӫ��a�̦h�����@���^�I"
location.href = "admin.asp"
</script>
<%
elseif jingao then
rs.close
conn.close
set conn=nothing
set rs=nothing
call showchat("<font color=ff0000>�iô�γ�ĵ�j</font>" & nickname & "�b�x���Ū��չϥH�@���覡��ID��" & id & "���޲z����Ϲﲼ�A�Qô���d�I")
%>
<script language=vbscript>
MsgBox "���n�o�˵L��n�ܡH�o���Oĵ�i�A�U�����A�n�ݡI"
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
MsgBox "�z�����Ť�������벼�I"
location.href = "admin.asp"
</script>
<%else
sql="SELECT piao FROM �Τ� WHERE ����='������' and ID=" & id
set rs=conn.execute(sql)
if rs.eof or rs.bof then
rs.close
conn.close
set conn=nothing
set rs=nothing
%>
<script language=vbscript>
MsgBox "�z��ְڡH�n���������̨S���o�ӤH�a�I"
location.href = "admin.asp"
</script>
<%else
sql="update �Τ� set piao=piao-1 where id=" & id
conn.execute sql
sql="update �Τ� set piao=piao+1 where id=" & myid
conn.execute sql
conn.close
set rs=nothing
set conn=nothing
%>
<script language=vbscript>
MsgBox "���±z�A�A�w�g���\����F�@�ӤϹﲼ�A��誺����v�|��@" & chr(10) & "�����|�U���ˬd����v���t�ƪ��޲z���ä��H�B�@�I"
location.href = "admin.asp"
</script>
<%
'end if
end if
end if
end if
%>
<!--#include file="chat.asp"-->