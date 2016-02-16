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
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
diqu=trim(request.form("diqu"))
if InStr(diqu,"or")<>0 or InStr(diqu,"'")<>0 or InStr(diqu,">")<>0 or InStr(diqu,"=")<>0 or InStr(diqu,"<")<>0 or InStr(diqu,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
	Response.End 
end if
nianling=trim(request.form("nianling"))
if InStr(nianling,"or")<>0 or InStr(nianling,"'")<>0 or InStr(nianling,">")<>0 or InStr(nianling,"=")<>0 or InStr(nianling,"<")<>0 or InStr(nianling,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
	Response.End 
end if
email=trim(request.form("email"))
if InStr(email,"or")<>0 or InStr(email,"'")<>0 or InStr(email,">")<>0 or InStr(email,"=")<>0 or InStr(email,"<")<>0 or InStr(email,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
	Response.End 
end if
oicq=trim(request.form("oicq"))
if InStr(oicq,"or")<>0 or InStr(oicq,"'")<>0 or InStr(oicq,">")<>0 or InStr(oicq,"=")<>0 or InStr(oicq,"<")<>0 or InStr(oicq,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
	Response.End 
end if
touxiang=trim(request.form("touxiang"))
if InStr(touxiang,"or")<>0 or InStr(touxiang,"'")<>0 or InStr(touxiang,">")<>0 or InStr(touxiang,"=")<>0 or InStr(touxiang,"<")<>0 or InStr(touxiang,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
	Response.End 
end if
qianming=trim(request.form("qianming"))
if len(qianming)>100 then 
	Response.Write "<script language=JavaScript>{alert('太多了，簽名要這麼多嗎，最多100個字？！');window.close();}</script>"
	Response.End 
end if
if InStr(qianming,"or")<>0 or InStr(qianming,"'")<>0 or InStr(qianming,">")<>0 or InStr(qianming,"=")<>0 or InStr(qianming,"<")<>0 or InStr(qianming,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
	Response.End 
end if
sql="update 用戶 set 地區='"&diqu&"',年齡='"&nianling&"',信箱='"&email&"',Oicq='"&oicq&"',頭像='"&touxiang&"',簽名='"&qianming&"' where 姓名='"&Name&"'"
conn.Execute(sql)
%>
<html>
<title>江湖資料修改</title>
<body background="../images/8.jpg">
<table border=1 align=center width=336 cellpadding="10" cellspacing="13" height="293" background="../images/12.jpg">
  <tr valign="top">
    <td background="../images/7.jpg" height="226"> 
      <div align="center">
<p>&nbsp;</p>
<p><font color="#000000" size="3">您成功修改資料!</font><font color="#FF3333" size="3"><br>
</font></p>
<p><font color="#FF3333" size="3"> </font> </p>
</div>
<table width=100%>
<tr>
<td>
<p align=center style='font-size:14;color:red'><font color="#000000"><br>
如果修改有錯誤，請重新修<br>
改即可！</font></p>
<p align=center>
<input type=button value=關閉窗口 onClick='window.close()' name="button">
</p>
</td>
</tr>
</table>
    </td>
</tr>
</table>
</html>
<html></html> 