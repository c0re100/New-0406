<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%response.expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
name=info(0)
face=left(trim(request.form("face")),3)
if InStr(face,"or")<>0 or InStr(face,"'")<>0 or InStr(face,"`")<>0 or InStr(face,"=")<>0 or InStr(face,"-")<>0 or InStr(face,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
Response.End 
end if
conn.Execute "update 用戶 set 名單頭像='"&face&"' where id="&info(9)
conn.close
set conn=nothing
%><html>
<head>
<title>更改頭像</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="<%=chatbgcolor%>" background="../1.jpg" bgproperties="fixed" leftmargin="0" topmargin="30">
<p><font size="+1" color="#FFFFFF">你已經成功的修改了頭像<img src="../ico/<%=face%>-2.gif"></font><font size="+1" color="#FF6633">退出江湖重新進入您就可以看到你更改後的頭像！</font></p>
<p><br>
</p>
</body>
</html> 
<html><script language="JavaScript">                                                                  </script></html>