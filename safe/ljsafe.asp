<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5> 
<title>修改密碼</title>
<style type="text/css">
</style>
<link rel="stylesheet" href="Setup.css">
</head>
<body oncontextmenu=self.event.returnValue=false text="#FFFFFF" LEFTMARGIN="0" TOPMARGIN="0" bgcolor="#990000">
<%
if not IsArray(Session("info")) then 
%>
<div align="center">對不起，您沒有登陸，申請密碼保護必需登陸後方可申請！ <br> <A HREF="#" ONCLICK="window.close()">[ 
關 閉 窗 口 ]</A><%
Response.End 
end if
info=Session("info")
%> </div><center> <p><font color="yellow">◇ 申請密碼保護 ◇</font></p><table border=1 bgcolor="#008080" align=center cellpadding="10" cellspacing="13" bordercolordark="#FFFFFF" bordercolorlight="#000000" width=380> 
<tr><td bgcolor=#008080 width="378"> <table bgcolor="#008080" border="1" align="center" width="100%" height="10" cellspacing="0" bordercolordark="#008080" bordercolorlight="#008080"> 
<tr> <td height="1"> <form method=POST action='ljsafeok.asp'> <div align="center"> 
<table width="100%" border="0" cellspacing="3" cellpadding="0" align="center"> 
<tr> <td>姓  名：</td><td> <input type=text name=name readonly size=13 maxlength="10" value=<%=info(0)%>> 
</td><td>登陸江湖的賬號</td></tr> <tr> <td>原 密 碼：</td><td> <input type=password name=oldpass size=13 maxlength="10"> 
</td><td>登陸江湖的密碼</td></tr> <tr> <td>找回密碼： </td><td> <input type=password name=getpass size=13 maxlength="10"> 
</td><td>要申請的密碼保護</td></tr> <tr> <td>校驗密碼：</td><td> <input type=password name=repass size=13 maxlength="10"> 
</td><td>要申請的密碼保護</td></tr> <tr> <td>郵  箱：</td><td> <input type=text name=mail size=13 maxlength="30"> 
</td><td>自己常用郵箱</td></tr> </table><br> </div><div align="center"> <input type=submit value="申 請" name="submit"> 
<input type=button value="關 閉" onClick="window.close()" name="button"> </div></form><p> 
申請的密碼保護可用來丟失密碼後初始化密碼！<br> 因此！請一定要記住您的找回密碼和郵箱！ </p></td></tr> </table></table></center>  
</body>  
</html> 
 
 
