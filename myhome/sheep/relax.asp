<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<!--#include file="data.asp"-->
<%
info=Session("info")
if info(0)="" then response.redirect"../../error.asp?id=210"
sql="SELECT * FROM �d�����A where user='"&info(0)&"' and �W�r='"&request("name")&"'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z���d���w�g���F�αz���O�o���d�����D�H�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
if rs("�ѧO")>=rs("����I") then 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���d���֤F�ӷ����F�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if
randomize
rad=Int(62 * Rnd + 20)
if rad>0 and rad<20 then rad=1
if rad>20 and rad<40 then rad=2
if rad>40 and rad<60 then rad=3
if rad>60 and rad<80 then rad=4
Select Case rad
case "1"
name=request("name")
sql="update �d�����A set �ѧO=�ѧO+1,��O=��O+50 where user='"&info(0)&"'"
conn.Execute(sql)
mess="�d���o��𮧡A�W�[��O50�I�C"
conn.close
case "2"
name=request("name")
sql="update �d�����A set �ѧO=�ѧO+1,��O=��O+100 where user='"&info(0)&"'"
conn.Execute(sql)
mess="�d���o��𮧡A�W�[��O100�I�C"
conn.close
case "3"
name=request("name")
sql="update �d�����A set �ѧO=�ѧO+1,��O=��O+200 where user='"&info(0)&"'"
conn.Execute(sql)
mess="�d���o��𮧡A�W�[��O200�I�C"
conn.close
case "4"
name=request("name")
sql="update �d�����A set �ѧO=�ѧO+1,��O=��O+150 where user='"&info(0)&"'"
conn.Execute(sql)
mess="�d���o��𮧡A�W�[��O150�I�C"
conn.close
Case else
name=request("name")
sql="update �d�����A set �ѧO=�ѧO+1,��O=��O+10 where user='"&info(0)&"'"
conn.Execute(sql)
mess="�d���o��𮧡A�W�[��O10�I�C "
conn.close
End Select
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�鮧</title>
<link rel="stylesheet" href="setup.css">
</head>

<body topmargin="0" leftmargin="0" bgcolor="#3a4b91" text="#000000" background="../../linjianww/0064.jpg">
<div align="center">
<center>
<br>
<br>
</center>
</div>
<div align="center">
<center>
<table border="1" width="378" cellspacing="1" cellpadding="0" height="173" bordercolor="#000000">
<tr>
<td class="p2" width="100%">
<div align="center"><font size="2" color="#000000"> <%=mess%><br>
<br>
</font> </div>
<table border="0" width="320" cellspacing="0" cellpadding="0" align="center">
<td class="p3" width="100%" height="19">
<p align="center"><font color="#000000"><a href="feedsheep.asp">&gt;&gt;��^</a>
</font>
</td>
</table>

</td>
</tr>
</table>
</center>
</div>
</body>

</html>