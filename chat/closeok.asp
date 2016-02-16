<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<title>關閉窗口</title>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor=FFFFFF>
<%Session.Abandon%>
<script LANGUAGE="JavaScript">
if(window!=window.top){top.location.href=location.href;}
window.close();
</script>
</body>
</html> 