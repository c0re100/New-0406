<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
grade=info(2)
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>在線交易</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
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
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" leftmargin="0" bgproperties="fixed" topmargin="0"  background=../bk.jpg>
<div align=center>
<table cellpadding="1" cellspacing="0" border="2" bgcolor="#EEFFEE" align="center" width=98% bordercolorlight="#000000" bordercolordark="#FFFFFF" bordercolor="#FFFFFF">
<tr align=center bgcolor=ffcc00><td>物品</td><td>數量</td><td>交易</td></tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT 物品名,數量,id FROM 物品 WHERE (類型<>'卡片') and 裝備<>true and 擁有者="&info(9)
rs.open sql,conn,1,1
do while not rs.eof and not rs.bof

%>
<tr valign="middle">
<td width="39%" height="15">
<div align="center"><font color="#000000" size="2" > <%=rs("物品名")%> </font></div>
</td>
<td width="26%" height="15"><font color="#000000" size="2" ><%=rs("數量")%></font></td>
<td width="35%" height="15">
<div align="center"><font color="#330033" size="2"> <a href="#" onClick="window.open('jiaoyi_wu.ASP?id=<%=rs("id")%>','tie','scrollbars=yes,resizable=yes,top=25,left=200,width=550,height=300')">出售</a></font></div>
</td>
</tr>
<%
rs.movenext
loop
rs.close
conn.close
set rs=nothing
set conn=nothing%>
</table>

</div>
</body>
</html>
