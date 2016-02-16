<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
nowdate=cstr(date())
bkmn=Request.Form("money")
bkop=Request.Form("op")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if isnumeric(bkmn) then
	if bkmn>1 then
		if bkop="取款" then	
conn.Execute "update 用戶 set 銀兩=銀兩+"&bkmn&",存款=存款-"&bkmn&" where 姓名='"&username&"' and 存款>="&bkmn
msg="您從錢莊裡成功取出<font color=red>["&bkmn&"]</font>兩銀子，哎，省著點花耶！"
end if
		if bkop="存款" then 
conn.Execute "update 用戶 set 銀兩=銀兩-"&bkmn&",存款=存款+"&bkmn&" where 姓名='"&username&"' and 銀兩>="&bkmn
msg="您在錢莊裡成功存入<font color=red>["&bkmn&"]</font>兩銀子，多多存錢啊，將來要養家糊口的啊！"
end if
set rs=nothing
conn.close
set conn=nothing		
		
	else
		msg="開玩笑的吧，你到底是想存還是想取！"
	end if	
else	
	msg="請準確輸入金額！"
end if	

%><head><link rel="stylesheet" href="../style.css"></head>

<title>靈劍江湖錢莊</title>
<body bgcolor='<%=bgcolor%>' background='../bg.gif' oncontextmenu=self.event.returnValue=false>
<p align=center>靈劍江湖錢莊 <div align="center"> <table width="75%" border="0" cellspacing="1" cellpadding="4" bgcolor="#000000"> 
<tr> <td bgcolor="#FFFFFF"> <p> <%=msg%> <p align=center><a href="bankok.asp" onMouseOver="window.status='返回錢莊';return true;" onMouseOut="window.status='';return true;"><b></b></a></p></td></tr> 
</table></div><p align="center"><a href="bankok.asp" onmouseover="window.status='返回錢莊';return true;" onmouseout="window.status='';return true;"><b>返回</b></a> 
<p align="center">&nbsp;
</body>