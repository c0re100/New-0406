<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="fun.inc"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if

name=info(0)
%>
<!--#include file="jhb.asp"-->
<%
	sql="select * from �Ȥ� where �b��='" & name & "'"
	set rs=conn.execute(sql)
	if rs.eof or rs.bof then
	set rs=nothing
conn.close
set conn=nothing%>
<script language="vbscript">
  alert "�A�٨S���}��O�I"
  location.href = "jingji.asp"
</script>
<%	   
    else
    jjr=rs("�g���H")
%>
<html>
<head>
<title>�g���H�줽��</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style>
td{color:#000000}
p{font-size:16;color:red}
</style>
</head>
<body bgcolor=#000000>
<table border=1 bgcolor="#BEE0FC" align=center width=350 cellpadding="10" cellspacing="13">
<tr><td bgcolor=#CCE8D6>
<table><tr><td>
<p align=center style="font-size:14;color:#000000"></p> 
</td></tr>  

<tr><td style="color:red;font-size:9pt">  
<br><p align=center>�z���Ѳ��g�٤H<%=jjr%>�b�����z�A��</p><br>
<a href=cun.asp>�s���i�Ѳ��b��</a><br><a href=qu.asp>�q�Ѳ��b�ᴣ��</a><br><a href=cha.asp>�d�ݳ̪�Ѳ��R�污�p</a><br><a href=huan.asp>���g���H</a>  
<br>  
<p align=center><a href=jingji.asp>���}</a></p>
</table></table>  

		
</body>
</html>
<%
end if
end if
%>