<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
toname=Trim(Request.QueryString("toname"))
if toid=0 then toid=info(9)
if toname="大家" then toname=info(0)
if info(2)<10 then
Response.Write "<script Language=Javascript>alert('你想作什麼滾！');window.close();</script>"
Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM 用戶 WHERE 姓名='"&toname&"'",conn
idd=rs("id")
rs.close
rs.open "SELECT * FROM 物品 WHERE 擁有者=" & idd & " and 數量>0  order by 類型 ",conn
%>
<html>
<head>
<title>物品管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: 新細明體; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px; BACKGROUND: #ffffff;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 新細明體; font-size: 9pt; font-style: bold;}
</style>
<style TYPE="text/css">
body{overflow:scroll;overflow-x:hidden}
.popper{  position : absolute;  visibility : hidden;}
</style>
<STYLE TYPE="text/css">
BODY {OVERFLOW:scroll;OVERFLOW-X:hidden}
.DEK {POSITION:absolute;VISIBILITY:hidden;Z-INDEX:200;}
</STYLE>
</head>
<body bgcolor="#660000"  background=../ff.jpg leftmargin="0" topmargin="0">
<div align="center"><%=toname%>現有物品一覽(裝備的物品除外)<a href="javascript:this.location.reload()">刷新</a></div>
<br>
<table border="1" align="center" width=80% cellpadding="1" cellspacing="0" height="31">
<tr align=center>
<td>物品名</td><td>類型</td><td>數量</td><td>內力</td><td>體力</td><td>攻擊</td><td>防御</td><td>等級</td><td>堅固度</td><td>價錢</td><td>裝備</td><td>方式</td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td><%=rs("物品名")%> </div></td>
<td><%=rs("類型")%></div></td>
<td><%=rs("數量")%></div></td>
<td><%=rs("內力")%></div></td>
<td><%=rs("體力")%></div></td>
<td><%=rs("攻擊")%></div></td>
<td><%=rs("防御")%></div></td>
<td><%=rs("等級")%></div></td>
<td><%=rs("堅固度")%></div></td>
<td><%=rs("銀兩")%></div></td>
<td><%=rs("裝備")%></div></td>
<td><a href="del.asp?id=<%=rs("id")%>">刪除</a><br><a href="decc.asp?id=<%=rs("id")%>">數量加10</a><br><a href="decc20.asp?id=<%=rs("id")%>">數量加20</a><br><a href="decc100.asp?id=<%=rs("id")%>">數量加100</a></td>
</tr>
<%
rs.movenext
loop
%>
</table>
<%
rs.close
rs.open "SELECT * FROM 交易市場 WHERE 擁有者=" & idd & " order by 方式",conn
%>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
<tr align="center">
<td nowrap width="45" height="16"><div align="center"> 物品名 </div></td>
<td nowrap width="40" height="16"><div align="center"> 類型  </div></td>
<td nowrap width="33" height="16"><div align="center"> 數量   </div></td>
<td nowrap width="40" height="16"><div align="center"> 內力  </div></td>
<td nowrap width="40" height="16"><div align="center"> 體力  </div></td>
<td nowrap width="40" height="16"><div align="center"> 攻擊  </div></td>
<td nowrap width="40" height="16"><div align="center"> 防御  </div></td>
<td colspan="2" nowrap height="16"><div align="center"> 價錢 </div></td>
<td nowrap width="45" height="16"><font size="-1">方式 </td>
<td nowrap width="50" height="16"><div align="center"> 操作 </div></td>
</tr>
<%
do while not rs.eof
%>
<tr>
<td width="45" height="3"><div align="center"> <%=rs("物品名")%>  </div></td>
<td width="40" height="3"><div align="center"> <%=rs("類型")%> </div></td>
<td width="33" height="3"><div align="center"> <%=rs("數量")%>  </div></td>
<td width="40" height="3"><div align="center"> <%=rs("內力")%> </div></td>
<td width="40" height="3"><div align="center"> <%=rs("體力")%> </div></td>
<td width="40" height="3"><div align="center"> <%=rs("攻擊")%> </div></td>
<td width="40" height="3"><div align="center"> <%=rs("防御")%> </div></td>
<td colspan="2" height="3"><div align="center"> <%=rs("銀兩")%>  </div></td>
<td height="3" width="45"><div align="center"> <%=rs("方式")%> </div></td>
<td height="3" width="50"><div align="center"><font size="-1"><a href="del1.asp?id=<%=rs("id")%>"><font color="#0000FF">刪除 </a> </div></td>
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
<tr align="center">
<td nowrap width="45" height="16">
<div align="center"></div>
</td>
</tr>
<table width="64%" border="1" cellpadding="0" cellspacing="0" align="center">
<tr>
<td height="25">
<div align="center">這裡可以查看到對方的物品，刪除就可以將此用戶的物品刪除！</div>
</td>
</tr>
</table>
</body>
</html>
