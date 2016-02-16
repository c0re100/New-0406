<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>
<head>
<title>寵物商店區</title>
<link rel="stylesheet" href="../../css.css">
</head>
<body leftmargin="0" topmargin="7" bgcolor="#FFFFFF">
<div align="center">
<div align="center">
<table width="778" border="0" cellspacing="0" cellpadding="4">
<tr>
<td valign="top" align="center"> [ <font color="#009900">寵 
物 商 店</font> ]<br> 
<br>
          【<a href="indexsheep.asp"> 寵物商店 </a>】 【<a href="stunt.asp"> 技能商店</a> 
          】 【 <a href="itemshop.asp">道具商店 </a>】 【 <a href="at.asp">寵物練武場</a> 】<br> 
<br> 
<div align="center"> </div> 
<div align="center"><font size="-1">歡迎光臨寵物商店，這裡有各式不同種類的寵供你選購呵。注意，寵物隻能買一隻呵。</font><br> 
<!--#include file="data.asp"--> <br> 
<br> 
<table width="572" border="1" cellspacing="2" cellpadding="0" align="center" bordercolor="eeeeee"> 
<tr> 
<td width="100%" colspan="4"> 
<div align="center">現 有 寵 物</div> 
</td> 
</tr> 
<tr> 
<td width="90"> 
<div align="center">寵物類型 </div> 
</td> 
<td width="222"> 
<div align="center">寵物參數 </div> 
</td> 
<td width="85"> 
<div align="center">售 價 </div> 
</td> 
<td width="59"> 
<div align="center">操 作 </div> 
</td> 
</tr> 
<% 
sql="SELECT 寵物類型,攻擊,防御,體力,價格 FROM 寵物'" 
Set rs=conn.Execute(sql) 
do while not rs.bof and not rs.eof 
%> 
<tr> 
<td width="90"><font color="#0000FF"><%=rs("寵物類型")%></font></td> 
<td width="222"> 
<div align="center"><font color="#0000FF">攻擊：<%=rs("攻擊")%> 
防御：<%=rs("防御")%> 體力：<%=rs("體力")%> </font></div> 
</td> 
<td width="85"> 
<div align="center"><%=rs("價格")%>兩</div> 
</td> 
<td width="59"> 
<p align="center"><font color="#0080FF"><a href="buy.asp?name=<%=rs("寵物類型")%>"><img border="0" src="goumai.gif" width="53" height="15"></a></font></p> 
</td> 
</tr> 
<% 
rs.movenext 
loop 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%> 
</table> 
</div> 
</td> 
</tr> 
</table> 
</div> 
</div> 
</body>

</html>