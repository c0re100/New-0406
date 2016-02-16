<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
action=request.querystring("action")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 物品名,數量,體力,內力,銀兩 FROM 物品 WHERE 擁有者=" & info(9) & " and 數量>0 and 類型='藥品'  order by 銀兩 ",conn
%>
<html>
<head>
<title>我的包袱</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body background="back.gif">
<div align="left"> <table width="453" border="0" cellspacing="0" cellpadding="0" height="224" align="center"> 
<tr> <td width="92" rowspan="3" valign="top"> <div align="center"><br> <input onClick="javascript:window.document.location.href='index.asp'" title=裝備一覽 type=button value="裝備一覽" name="button7"> 
<br> <br> <input onClick="javascript:window.document.location.href='head.asp'" title=武器裝備 type=button value="武器裝備" name="button"> 
<br> <br> <input onClick="javascript:window.document.location.href='eat.asp'" title=喫用藥物 type=button value="喫用藥物" name="button62"> 
</div></td><td valign="top" rowspan="3" width="361"> <div align="center"><img src="dao.gif" width="40" height="15">喫用藥品一覽<img src="dao1.gif" width="40" height="15"> 
<font color="#CC0000" face="幼圓"><a href="javascript:this.location.reload()">刷新</a></font></div><div align="center"> 
注意:選擇服用，將服用所有的藥品！<br> <br> <table border="1" align="center" width="326" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000"> 
<tr align="center"> <td nowrap width="65" height="15"> <div align="center"><font color="#000000">物品名</font></div></td><td nowrap width="30" height="15"> 
<div align="center"><font color="#000000">數量 </font> </div><td nowrap width="36" height="15"> 
<div align="center"><font color="#000000">加體力</font></div></td><td nowrap width="36" height="15"> 
<div align="center"><font color="#000000">加內力</font></div></td><td colspan="2" nowrap height="15"> 
<div align="center"><font color="#000000">價值</font></div></td><td width="49" height="15"> 
<div align="center"><font color="#000000">使用數量</font></div></td><td width="63" height="15"> 
<div align="center"><font color="#000000">操作</font></div></td></tr> <%
do while not rs.eof
%> <tr> <form method=POST action='wupin1.asp?action=<%=rs("物品名")%>&name=<%=info(0)%>'> 
<td width="65" height="10"><div align="center"><%=rs("物品名")%> </div></td><td width="30" height="10"> 
<div align="center"><font color="#FF0000"><%=rs("數量")%> </font> </div><td width="36" height="10"> 
<div align="center"><font color="#0000FF"><%=rs("體力")%></font> </div></td><td width="36" height="10"> 
<div align="center"><font color="#0000FF"><%=rs("內力")%> </font> </div></td><td colspan="2" height="10"> 
<div align="center"><font color="#FF0000"><%=rs("銀兩")%> </font> </div></td><td width="49" height="10"> 
<div align="center"> <select name="wpsl"> <option value="1" selected>1</option> 
<option value="2">2</option> <option value="3">3</option> <option value="4">4</option> 
<option value="10">10</option> <option value="15">15</option> <option value="20">20</option> 
<option value="25">25</option> <option value="50">50</option> </select> </div></td><td width="63" height="10"> 
<div align="center"> <input type="SUBMIT" name="Submit2"  value="進行"> </div></td></form></tr> 
<%
rs.movenext
loop
%> </table><br> <%
rs.close
rs.open "SELECT 物品名,數量,功效,增加數值 FROM 修練物品 WHERE 擁有者=" & info(9) & " and 數量>0  order by 增加數值 ",conn
%> 修練藥品列表<br> <table border="1" align="center" width="326" cellpadding="0" cellspacing="0" height="23" bordercolor="#000000"> 
<tr align="center"> <td nowrap width="65" height="15"> <div align="center"><font color="#000000">物品名</font></div></td><td nowrap width="30" height="15"> 
<div align="center"><font color="#000000">數量 </font> </div><td nowrap width="36" height="15"> 
<div align="center"><font color="#000000">功效</font></div></td><td colspan="2" nowrap height="15"> 
<div align="center"><font color="#000000">價值</font></div></td><td width="49" height="15"> 
<div align="center"><font color="#000000">使用數量</font></div></td><td width="63" height="15"> 
<div align="center"><font color="#000000">操作</font></div></td></tr> <%
do while not rs.eof
%> <tr> <form method=POST action='xleat.asp?action=<%=rs("物品名")%>&name=<%=name%>'> 
<td width="65" height="17"><div align="center"><%=rs("物品名")%> </div></td><td width="30" height="17"> 
<div align="center"><font color="#FF0000"><%=rs("數量")%> </font> </div><td width="36" height="17"> 
<div align="center"><font color="#0000FF"><%=rs("功效")%> </font> </div></td><td colspan="2" height="17"> 
<div align="center"><font color="#FF0000"><%=rs("增加數值")%> </font> </div></td><td width="49" height="17"> 
<div align="center"> <select name="xlsl"> <option value="1" selected>1</option> 
<option value="2">2</option> <option value="3">3</option> <option value="4">4</option> 
<option value="5">5</option> <option value="6">6</option> <option value="7">7</option> 
<option value="8">8</option> <option value="9">9</option> </select> </div></td><td width="63" height="17"> 
<div align="center"> <input type="SUBMIT" name="Submit22"  value="進行"> </div></td></form></tr> 
<%
rs.movenext
loop

%> </table><BR><%
rs.close
rs.open "SELECT 物品名,數量,內力,銀兩 FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='鮮花' and 數量>0  order by 銀兩 ",conn
%> 鮮花列表<BR><TABLE BORDER="1" ALIGN="center" WIDTH="326" CELLPADDING="0" CELLSPACING="0" HEIGHT="23" BORDERCOLOR="#000000"> 
<TR ALIGN="center"> <TD nowrap WIDTH="65" HEIGHT="15"> <DIV ALIGN="center"><FONT COLOR="#000000">物品名</FONT></DIV></TD><TD nowrap WIDTH="30" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">數量 </FONT> </DIV><TD nowrap WIDTH="36" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">功效</FONT></DIV></TD><TD COLSPAN="2" nowrap HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">價值</FONT></DIV></TD><TD WIDTH="49" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">使用數量</FONT></DIV></TD><TD WIDTH="63" HEIGHT="15"> 
<DIV ALIGN="center"><FONT COLOR="#000000">操作</FONT></DIV></TD></TR> <%
do while not rs.eof
%> <TR> <FORM METHOD=POST ACTION='eathua.asp?action=<%=rs("物品名")%>&name=<%=name%>'> 
<TD WIDTH="65" HEIGHT="17"><DIV ALIGN="center"><%=rs("物品名")%> </DIV></TD><TD WIDTH="30" HEIGHT="17"> 
<DIV ALIGN="center"><FONT COLOR="#FF0000"><%=rs("數量")%> </FONT> </DIV><TD WIDTH="36" HEIGHT="17"> 
<DIV ALIGN="center"><FONT COLOR="#0000FF"><%=rs("內力")%> </FONT> </DIV></TD><TD COLSPAN="2" HEIGHT="17"> 
<DIV ALIGN="center"><FONT COLOR="#FF0000"><%=rs("銀兩")%> </FONT> </DIV></TD><TD WIDTH="49" HEIGHT="17"> 
<DIV ALIGN="center"> <SELECT NAME="hua"> <OPTION VALUE="1" selected>1</OPTION> 
<OPTION VALUE="2">2</OPTION> <OPTION VALUE="3">3</OPTION> <OPTION VALUE="4">4</OPTION> 
<OPTION VALUE="5">5</OPTION> <OPTION VALUE="6">6</OPTION> <OPTION VALUE="7">7</OPTION> 
<OPTION VALUE="8">8</OPTION> <OPTION VALUE="9">9</OPTION> </SELECT> </DIV></TD><TD WIDTH="63" HEIGHT="17"> 
<DIV ALIGN="center"> <INPUT TYPE="SUBMIT" NAME="Submit222"  VALUE="進行"> </DIV></TD></FORM></TR> 
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%> </TABLE><br> <input onClick="JavaScript:window.close()" title=關閉窗口 type=button value="關閉窗口" name="button2"> 
</div></td></tr> <tr> </tr> <tr> </tr> </table></div>
</body>
</html>
