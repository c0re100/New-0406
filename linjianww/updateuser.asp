<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
userid=Request.Form("id")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
Set rst=Server.CreateObject("ADODB.RecordSet")
sqlstr="SELECT * FROM �Τ� where id="&userid
rst.Open sqlstr,conn,3,3
if Request.Form("submit")="�R��" then
if info(0)<>"�����`��" then
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
name=rst("�m�W")
sqlstr="delete * from �Τ� where id="&userid
conn.Execute sqlstr
sqlstr="delete * from ���~ where �֦���="&userid
conn.Execute sqlstr
sqlstr="delete * from ������� where �֦���="&userid
conn.Execute sqlstr
sqlstr="delete * from �ޯ�C�� where �֦���='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from �歱 where �֦���='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from �ϥΧޯ� where �m�W='"&name&"'"
conn.Execute sqlstr
call isin(name)
Response.Write "���\�R��id="&userid&"���Τ�"
elseif Request.Form("submit")="�R���L�����~" then
name=rst("�m�W")
sqlstr="delete * from ���~ where �֦���="&userid
conn.Execute sqlstr
sqlstr="delete * from ������� where �֦���="&userid
conn.Execute sqlstr
sqlstr="delete * from �ޯ�C�� where �֦���='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from �歱 where �֦���='"&name&"'"
conn.Execute sqlstr
sqlstr="delete * from �ϥΧޯ� where �m�W='"&name&"'"
conn.Execute sqlstr
call isin(name)
Response.Write "���\�R��id="&userid&"���Τ᪺���~"
else
for i=1 to rst.Fields.Count-1
Response.Write rst.Fields(i).Name&"("&rst.Fields(i).Type&")"&rst.Fields(i).Value&"---->"&Request.Form(i+1)&"<br>"
if rst.Fields(i).type=202 then
rst.Fields(i).Value=cstr(Request.Form(i+1))
elseif rst.Fields(i).type=3 then
rst.Fields(i).Value=clng(Request.Form(i+1))
elseif rst.Fields(i).type=6 then
rst.Fields(i).Value=Request.Form(i+1)
elseif rst.Fields(i).Type=135 then
rst.Fields(i).Value=cdate(Request.Form(i+1))
end if	
next
rst.Update
end if
rst.Close
set rst=nothing
conn.Close
set conn=nothing
function isin(name)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
gupiao="DBQ="+server.mappath("../gupiao/gp.asp")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open gupiao
'����Τ�
sql="SELECT * FROM �j�� WHERE �b��='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "delete * from �j�� where �b��='"&name&"'"
end if
sql="SELECT * FROM �Ȥ� WHERE �b��='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "delete * from  �Ȥ�  where �b��='"&name&"'"
end if
'�O��
end function
%>
<body background="0064.jpg">
<a href="fine.asp">��^</a> 