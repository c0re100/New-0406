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
a=request("a")
id=request("id")
name=ljjh_nickname
name2=request.form("name2")
if a="a" then
sql="SELECT * FROM �Τ� WHERE �m�W='" & name2 & "'"
else
sql="SELECT * FROM �Τ� WHERE ID=" & id
end if
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess=name2 & "���O���򤤤H�I"
else
sql="select * from �Τ� where �m�W='" & name & "' and ����='�x��' and ����='�x��'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess=name & "�A���O�x��"
else
'		if a="a" then
select case a
case "a"
sql="update �Τ� set grade=6, ����='�x��' where �m�W='" & name2 & "'"
conn.execute sql
mess="�A���\�a��" & name2 & "�۸u���x�����u�@�H��"
case "b"
sql="update �Τ� set ����='�̤l', ����='�C�L' where id=" & id
conn.execute sql
mess="�A���\�a��" & id & "�q�x���}���F�I"
case "c"
sf=request("sf")
sql="update �Τ� set grade='" & sf & "' where id=" & id
conn.execute sql
mess="�A���\�a��" & id & "�������ɡ]���^�Ŭ�"&sf&"�F�I"
end select
end if
end if
conn.close
set conn=nothing
%>

<head>
<style>td           { font-size: 14 }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����</title></head>

<body background="../images/8.jpg">
<table border="1" align="center" width="350" cellpadding="10"
cellspacing="13" background="../images/12.jpg">
  <tr>
    <td height="89" background="../images/7.jpg"> 
      <table width="100%">
<tr>
<td>
<p align="center" style="font-size:14;color:red"><%=mess%></p>
<p align="center"><a href="admin.asp">��^</a></p>
</td>
</tr>
</table>
    </td>
</tr>
</table>
<div align="center"> <br>
  <font color="#FFFFFF" size="2">���F�C�����v�Ҧ��� </font></div>

</body>