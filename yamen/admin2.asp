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
id=request("id")
sql="SELECT * FROM �Τ� WHERE ��O<-1000 and ���A='��' and ����='�x��' and ID=" & id
set rs=conn.execute(sql)
if rs.eof or rs.bof then
mess="���H�����άO�۱��Τ��ۧA�h�ߡI"
else
sql="update �Τ� set ���A=' ', ��O=0 where id=" & id
conn.execute sql
mess="���\�a��ID" & id & "�Ϭ�"
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
<title>�F�C�����v�Ҧ�</title></head>

<body background="../images/8.jpg">
<table border="1" align="center" width="350" cellpadding="10"
cellspacing="13" height="134" background="../images/12.jpg">
  <tr>
    <td background="../images/7.jpg"> 
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
<div align="center">
<br>
  <font color="#FFFFFF" size="2">���F�C�����v�Ҧ���</font></div>

</body>
<html></html> 