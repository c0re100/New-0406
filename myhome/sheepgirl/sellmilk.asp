<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data2.asp"-->
<%
if info(0)="" then%>
<%else
sheepname=request.form("sheepname")
milk=request.form("milk")
if sheepname="" or not isnumeric(milk) then
%>
<script language="Vbscript">
msgbox"��J���~�I",0,"ĵ�i" 
history.back
</script>
<%
else
intmilk=milk-int(milk)
if intmilk<>0 or milk<1 then
%>
<script language="Vbscript">
msgbox"��J���~�I",0,"ĵ�i" 
history.back
</script>
<%else
rs.open"select * from sheep where user='"&info(0)&"' and sheepname='"&sheepname&"'",conn,1,1
if rs.bof then
rs.close
conn.close
response.redirect"warning.htm"
else
milklast=rs("milk")-milk
if milklast<0 then
rs.close
conn.close%>
<script language="Vbscript">
msgbox"�A���p�ĨS�����o��h���u�@���ڡI",0,"ĵ�i" 
history.back
</script><%
else
rs.close
conn.execute"update sheep set milk='"&milklast&"' where user='"&info(0)&"' and sheepname='"&sheepname&"'"
rs.open"select * from rules",conn,1,1
milkprice=rs("milkprice")
rs.close
tempsplosh=milkprice*milk

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute"update �Τ� set �w���I��=�w���I��+'"&tempsplosh&"' where id="&info(9)

	set rs=nothing	
	conn.close
	set conn=nothing
%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����V����t��|</title>
<link rel="StyleSheet" href="new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
</div>
<div align="center">
  <center>
  <table border="0" width="470" bordercolor="#FFFFFF" cellspacing="1"
  cellpadding="0" height="20">
    <tr>
      <td class="p6"><font size="2">��-�z��e����m-&gt;����t��|�r&gt;��u��&nbsp;&nbsp;&nbsp; 
        <a href="javascript:history.back(1)">��^</a></font></td>
    </tr>
  </table>
  </center>
</div>
<div align="center">
  <center>
  <table border="0" width="470" cellspacing="1" cellpadding="0" height="50">
    <tr>
      <td class="p2" width="100%">
        <p align="center"><font size="2">���ߡI 
        �z�w�g���\��F�A���Ĥl���u��A�@��o���l�G<%=tempsplosh%>��</font></p>
      </td>
    </tr>
  </table>
  </center>
</div>

</body>

</html>

<% 
end if 
end if 
end if 
end if 
end if 
%>