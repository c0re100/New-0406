<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head>
<title>�c�Э��a�A�~�H���J</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body leftmargin="0" topmargin="0" background="../001/tanpopo7.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF">
<br>
<br>
<br>
<br>
<br>
<br>
<img src="../001/koi.gif" width="461" height="251" usemap="#Map" border="0"> 
<map name="Map">
  <area shape="rect" coords="127,48,214,71" a href ="listlao.asp" onClick="window.open('listlao.asp','yamen','scrollbars=no,resizable=no,top=25,left=100,width=570,height=400')" alt="�d�ݦb��H��" title="�d�ݦb��H��" > 
  <area shape="rect" coords="185,216,261,243" a href ="#" onClick="window.close()" alt="�h�X�Ū�" title="�h�X�Ū�">
</map>
<script language=javascript> 
     function Click(){
     alert('�F�C�����w��z!!!');
     window.event.returnValue=false;
     }
     document.oncontextmenu=Click;
     </script>

</body>

</html>

