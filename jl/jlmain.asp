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
<title>�b��H��</title>
<link rel="stylesheet" type="text/css" href="style.css">
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body leftmargin="0" topmargin="0" background="bk.jpg" link="#0000FF" vlink="#0000FF" alink="#0000FF" marginwidth="0" marginheight="0">
<DIV ID="overDiv" STYLE="position:absolute; visibility:hide; z-index: 1;; width: 179px; height: 16px"></DIV><SCRIPT LANGUAGE="JavaScript" src="../ljimage/lj1.js"></SCRIPT>
<img src="../jhimage/wj04.gif" width="510" height="383" usemap="#Map" border="0"> 
<map name="Map"> 
  <area shape="rect" coords="304,250,391,299" a href ="#" onClick="window.close()" ONMOUSEOVER="drs('::�h�X::');return true;" ONMOUSEOUT="nd(); return true;"> 
  <area shape="rect" coords="186,181,217,206" a href ="jl12.asp"   ONMOUSEOVER="drs('::�d�ݦ���ާ@�O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="178,256,209,281" a href ="jl11.asp"   ONMOUSEOVER="drs('::�d�ݦ����H�O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="172,228,203,253" a href ="jl10.asp"   ONMOUSEOVER="drs('::�d�ݦ����I�ްO��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="144,214,169,236" a href ="jl9.asp"    ONMOUSEOVER="drs('::�d�ݦ���ǰe�g��O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="218,174,243,196" a href ="jl8.asp"    ONMOUSEOVER="drs('::�d�ݦ���U�ʰO��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="245,175,261,204" a href ="jl7.asp"    ONMOUSEOVER="drs('::�d�ݦ�����y�ާ@�O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="261,158,277,181" a href ="jl6.asp"    ONMOUSEOVER="drs('::�d�ݦ���Ѱ��ʸT�O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="278,156,297,192" a href ="jl5.asp"    ONMOUSEOVER="drs('::�d�ݦ���ʸT�O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="197,282,223,308" a href ="jl4.asp"    ONMOUSEOVER="drs('::�d�ݦ��򧤨c�O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="296,162,322,188" a href ="jl3.asp"    ONMOUSEOVER="drs('::�d�ݦ���e���O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="329,194,355,220" a href ="jl2.asp"    ONMOUSEOVER="drs('::�d�ݦ���ĵ�i�O��::');return true;" ONMOUSEOUT="nd(); return true;">
  <area shape="rect" coords="364,207,390,233" a href ="jl1.asp"    ONMOUSEOVER="drs('::�d�ݦ���R�װO��::');return true;" ONMOUSEOUT="nd(); return true;">
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