<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select hysee from 栋 where id>=0 order by id",conn
hy=rs("hysee") 
if hy=1 then
rs.close
rs.open "select 穦单 from ノめ where id="&info(9),conn
hy1=rs("穦单")
if hy1=0 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('ぃ琌穦叫');location.href = '../help/info.asp?type=2&name=穦快猭';}</script>"
Response.End
end if
end if
%>
<html>
<head>
<title>穦琩高祘</title>
<style type="text/css">
<!--
p            { line-height: 20px; font-size: 9pt }
table        { font-size: 9pt }
a:link       { color: #FF9900; text-decoration: none }
-->
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#FFFFFF" vlink="#FF9900" topmargin="0"
leftmargin="0" bgcolor="#006666">
<p align="center"> <font color="#CC0000" face="ギ蛾"><a href="javascript:this.location.reload()">穝</a></font>
<br>
稰谅硂ㄇ狟ね癸и打や穦翴<a href="../help/info.asp?type=2&name=穦快猭" target="_blank">硂柑</a><br>

<table border="0" width="500" cellspacing="0" cellpadding="0"
background="bg.gif" align="center">
<tr align="center">
    <td background="top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1" color="#FFFF99"> 打 穦 </font></b></font></td>
</tr>
<tr align="center">
<td>
<table width="470" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF" height="13">
<tr>
<%rs.close
rs.open "SELECT ID,﹎,┦,,ō,穦单,穦ら戳,单,穦禣 FROM ノめ where 穦单>0 order by 穦ら戳 ",conn

%>
<td width="28" height="18">
<div align="center"><font color="#FFFFFF">ID</font></div>
</td>
<td width="47" height="18">
<div align="center"><font color="#FFFFFF">﹎</font></div>
</td>
<td width="25" height="18">
<div align="center"><font color="#FFFFFF">┦</font></div>
</td>
<td width="63" height="18">
<div align="center"><font color="#FFFFFF"></font></div>
</td>
<td width="86" height="18">
<div align="center"><font color="#FFFFFF">ō</font></div>
</td>
<td width="40" height="18">
<div align="center"><font color="#FFFFFF">穦</font></div>
</td>
<td width="75" height="18">
<div align="center"><font color="#FFFFFF">穦挡ら戳</font></div>
</td>
<td width="51" height="18">
<div align="center"><font color="#FFFFFF">打单</font></div>
</td>
<td width="35" height="18">
<div align="center"><font color="#FFFFFF">穦禣</font></div>
</td>
<td width="35" height="18">
<div align="center"><font color="#FFFFFF">э</font></div>
</td>
</tr>
<%
dengji1=0
dengji2=0
dengji3=0
dengji4=0
dengji5=0
dengji6=0
dengji7=0
dengji8=0
dengji9=0
dengji10=0
do while not rs.bof and not rs.eof
Select Case rs("穦单")
Case 1
	dengji1=dengji1+1
case 2
	dengji2=dengji2+1
case 3
	dengji3=dengji3+1
case 4
	dengji4=dengji4+1
case 5
	dengji5=dengji5+1
case 6
	dengji6=dengji6+1
case 7
	dengji7=dengji7+1
case 8
	dengji8=dengji8+1
case 9
	dengji9=dengji9+1
case 10
	dengji10=dengji10+1
end select
%>
<tr>
<td width="28" height="30">
            <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
</td>
<td width="47" height="30">
<div align="center"><a href="../yamen/mt.asp?action=<%=rs("﹎")%>"><font color="#FF9900"><%=rs("﹎")%></font></a></div>
</td>
<td width="25" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("┦")%></font></div>
</td>
<td width="63" height="30">
            <div align="center"><font color="#FFFFFF"><%=rs("")%></font></div>
</td>
<td width="86" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("ō")%></font></div>
</td>
<td width="40" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("穦单")%></font></div>
</td>
<td width="75" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("穦ら戳")%></font></div>
</td>
<td width="51" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("单")%></font></div>
</td>
<td width="35" height="30">
<div align="center"><font color="#FFFFFF"><%=rs("穦禣")%></font></div>
</td>
<td width="35" height="30">
<div align="center"><font color="#FFFFFF"><a href=./cc26.asp?id=<%=rs("ID")%>>ok</a></font></div>
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
</td>
</tr>
</table>
<div align="center"><font color="#000000">ノめ</font><b><font color="#0000FF"><%=dengji1%></font></b>
<font color="#000000">穦:</font><b><font color="#0000FF"><%=dengji2%></font></b><font color="#000000">
穦</font><font color="#0000FF"><b><%=dengji3%></b></font><font color="#000000">
穦</font><font color="#0000FF"><b><%=dengji4%></b></font><font color="#000000">
き穦</font><font color="#0000FF"><b><%=dengji5%></b></font><font color="#000000">
せ穦</font><font color="#0000FF"><b><%=dengji6%></b></font><font color="#000000">
穦</font><font color="#0000FF"><b><%=dengji7%></b></font><font color="#000000">
穦</font><font color="#0000FF"><b><%=dengji8%></b></font><font color="#000000">
穦</font><font color="#0000FF"><b><%=dengji9%></b></font><font color="#000000">
穦</font><font color="#0000FF"><b><%=dengji10%></b></font><font color="#000000"><br>
穦羆计</font><b><font color="#0000FF"><%=(dengji1+dengji2+dengji3+dengji4+dengji5+dengji6+dengji7+dengji8+dengji9+dengji10)%></font></b><font color="#000000"></font>
</div>
</body>
</html>