<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name1=Request.form("name1")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name2")
name2=trim(name2)
name2=server.HTMLEncode(name2)
kill=request.form("kill")
kill=trim(kill)
kill=server.HTMLEncode(kill)
mess=Request.form("mess")
mess=trim(mess)
money=Request.form("money")
if not isnumeric(money) then response.redirect"../error.asp?id=464"
money=int(money)
if money<1000000 or money>10000000000 then 
%>
<script language=vbscript>
MsgBox "���Ī�����W�L100���M�p��100�U"
location.href = "shashoulist.asp"
</script>
<%response.end
end if
if len(mess)>10 then 
%>
<script language=vbscript>
MsgBox "�n�����ܤ���W�L10�Ӧr��"
location.href = "shashoulist.asp"
</script>
<%response.end
end if
mess=server.HTMLEncode(mess)
if name1="" or name2="" or kill="" then
	Response.Write "<script language=JavaScript>{alert('�U���`�N��g�ɶq���n����!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
if name1=name2 or name1=kill or name2=kill or name1<>info(0) then
Response.Write "<script language=JavaScript>{alert('�A���m�W�B��������B�Q�����H�W�r����ۦP!!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'����Τ�
sql="SELECT ���� FROM �Τ� WHERE �m�W='" & name2 & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�S�o�ӤH�ڡH�d���F�a�I!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
if rs("����")<>"���" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�n���Ī����Ⱖ��O��Ȳ�´���H�H�I!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
rs.close
set rs=nothing

sql="SELECT ����,����,�Ȩ� FROM �Τ� WHERE �m�W='" & info(0) & "'"
Set Rs=conn.Execute(sql)
if rs("����")<3 or rs("����")="�x��" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�����ŤӧC�F�A�n�԰����Ťj��3�ũM�D�x�����ץi�H���ı���H�I!');location.href = 'gushashou.htm';}</script>"
	response.end
end if
if rs("�Ȩ�")<money then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A������h���ܡH�I!');location.href = 'gushashou.htm';}</script>"
response.end
end if
sql="update �Τ� set �Ȩ�=�Ȩ�-"&money&" where id="&info(9)
conn.execute sql
rs.close
sql="SELECT id FROM �Τ� WHERE �m�W='" & kill & "'"
Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('���򤤨S���A�n�����H�ڡH�d���F�a�I!');location.href = 'gushashou.htm';}</script>"
	response.end
end if

sz = "'" & name1 & "','" & kill & "','" & name2 & "','" & mess & "'," & money & ",now()"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
into_db = "INSERT INTO ���� (�m�W1,�Q����,�m�W2,����,���H�Ī�,�ɶ�) VALUES(" & sz & ")"
conn.Execute(into_db)
sql="delete * from ���� where now()-�ɶ�>10"
conn.execute sql
msg="<font color=FFD7F4>�p�D�����G</font><font color=FFD7F4>"&info(0)&"</font>���G<font color=FFD7F4>"&mess&"</font>�A��O�n��<font color=FFD7F4>"&kill&"</font>������<font color=FFD7F4>"&money&"</font>��Ȥl���Ĩ�Ȳ�´������<font color=FFD7F4>"&name2&"</font>,���F<font color=FFD7F4>"&kill&"</font>��<font color=FFD7F4>"&name2&"</font>�N�i�H�h����F�A�����I"

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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>"& msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

	Response.Write "<script language=JavaScript>{alert('�n�O���\�I!');location.href = 'shashoulist.asp';}</script>"
	response.end%>