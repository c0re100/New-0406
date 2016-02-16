<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
id=Trim(Request.QueryString("id"))
%><html>
<head>
<title>出錯提示</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body background=../images/8.jpg leftmargin="0" topmargin="0">
<table width="100%" border="0" height="100%">
<tr align="center">
<td>
<form method="post" action="">
        <table border="1" bordercolorlight="000000" bordercolordark="FFFFFF" cellspacing="0" background="../images/7.jpg">
          <tr>
<td>
<table border="0" bgcolor="#0000FF" cellspacing="0" cellpadding="2" width="350">
<tr>
<td width="360" align="center"><font color="#FFFFFF" size="3">出錯提示</font></td>
</tr>
</table>
<table border="0" width="350" cellpadding="4">
<tr>
<td width="328" valign="top">
<%Select Case id
Case "000"%>
<p>  程序所在目錄不是虛擬目錄，或設置有錯誤，Global.asa 沒有被執行。本程序需要虛擬目錄的支持！</p>
<%Case "001"%>
<p>  <font color=red><%=Request.QueryString("kickname")%></font> 並不在聊天室中，自救程序無法完成，查看是否房間選擇錯誤！</p>
<%Case "002"%>
<p>  是不是想開玩了？連大名和口令都不報？找死呀！</p>
<%Case "003"%>
<p>  是我不行，還是你不行？要找的人哪裡有呀？笨蛋！</p>
<%End Select%>
</td>
</tr>
<tr>
<td align="center" valign="top">
<input type="button" name="ok" value=" 確 定 " onclick=javascript:history.go(-1)>
</td>
</tr>
</table>
</td>
</tr>
</table>
</form>
</td>
</tr>
</table>
<script Language="JavaScript">
document.forms[0].ok.focus();
</script>
</body>
</html>
<html><script language="JavaScript">                                                                  </script></html> 