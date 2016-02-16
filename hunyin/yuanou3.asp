<%@ LANGUAGE = VBScript%>
<%Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>離婚</title>

<STYLE type=text/css>A:hover {
	CURSOR: resize
}
A {
	TEXT-DECORATION: none
}
select       { background-color: #FFFFFF; font-size: 9pt; border-left: medium solid #999900; 
              border-right: medium solid #FFCC66; 
               border-top: medium solid #999900; 
               border-bottom: medium solid #FFCC66  }
</STYLE>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<body text="#00FFFF" bgcolor="#FFFFFF" link="#99FF33" vlink="#FFFFFF" alink="#FF0000" leftmargin="0" topmargin="0">
<table width="691" border="0" cellspacing="0" cellpadding="0" height="119">
  <tr>
    <td background="lj1.gif"> 
      <p><img src="111.gif" width="77" height="111" align="left"><font color="#9900CC">你真的想清楚了嗎？真的要離開那個普經深愛過的人嗎？</font></p>
      <p><a href="yuanou4.asp">長痛不如短痛,還是離了</a> <a href="#" onClick="window.close()">考慮一下</a> 
      </p>
    </td>
  </tr>
</table>
<div align="center"></div>

</body>

</html>

 