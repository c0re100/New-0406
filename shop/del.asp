<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10  then  response.redirect "../error.asp?id=425"
id=Request.QueryString("id")
%><!--#include file="data.asp"--><%
sql="SELECT * FROM �ө� where �ө��W='"&id&"'"
rs1.open sql,conn1,1,1
if rs1.EOF or rs1.BOF then Response.Redirect "../error.asp?id=484"
id1=rs1("id")
mysql="delete * from �ө� where �ө��W='"&id&"'"
conn1.Execute (mysql)
mysql="delete * from �ө����~ where �֦���="&id1
conn1.Execute (mysql)
set rs1=nothing
conn1.close
set conn1=nothing
%>
<script language="vbscript">
alert("�R�����\�I")
location.href = "javascript:history.go(-1)"
</script>
