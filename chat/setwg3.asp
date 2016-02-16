<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
pai=info(5)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title><%=pai%>武學</title>
<script language="JavaScript">
function s(list){
var lijigongji=liji.checked;
parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();
if (lijigongji==true) {parent.f2.document.af.subsay.click()};
}
</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgproperties="fixed" topmargin="0" leftmargin="0" background=bk.jpg>
<div align="center">
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >
<tr><td width="53%" height="20" bgcolor="#FF6633">
<div align="center"><font color="#330033" size="2">招式名</font></div></td>
<td width="24%" height="20" bgcolor="#FF6633"><div align="center">等級</div></td>
<td width="23%" height="20" bgcolor="#FF6633"><font color="#330033" size="2">內力</font></td>
</tr>
<tr bgcolor=EEFFEE><td colspan="3" height="20"><div align="center">攻擊術</div></td></tr>
<%
rs.open "SELECT 武功,內力,級別 FROM 武功 where 門派='" & pai & "' and 類型='攻擊'",conn
do while not rs.eof and not rs.bof
%>
<tr  bgcolor="#EEFFEE">
<td height="20"><div align="center"><a href="javascript:s('/發招$ <%=rs("武功")%>')">
<input type=hidden name=wg1 value='<%=rs("武功")%>'><%=rs("武功")%></a></div></td>
<td height="20"><%=rs("級別")%></td><td height="20"><%=rs("內力")%></td>
</tr>
<%rs.movenext
loop
rs.close%>
<tr  bgcolor="#EEFFEE">
<td colspan="3" height="20">
<div align="center">防御術</div>
</td>
</tr>
<%
rs.open "SELECT 武功,內力,級別 FROM 武功 where 門派='" & pai & "' and 類型='防御'",conn
do while not rs.eof and not rs.bof
%>
<tr  bgcolor="#EEFFEE">
<td height="20">
<div align="center"><a href="javascript:s('/發招$ <%=rs("武功")%>')">
<input type=hidden name=wg1 value='<%=rs("武功")%>'>
<%=rs("武功")%></a></div>
</td>
<td height="20"><%=rs("級別")%></td>
<td height="20"><%=rs("內力")%></td>
</tr>
<%
rs.movenext
loop
rs.close%>
<tr  bgcolor="#EEFFEE">
<td colspan="3" height="20">
<div align="center">恢復術</div>
</td>
</tr>
<%rs.open "SELECT 武功,內力,級別 FROM 武功 where 門派='" & pai & "' and 類型='恢復'",conn
do while not rs.eof and not rs.bof
%>
<tr  bgcolor="#EEFFEE">
<td height="20">
<div align="center"><a href="javascript:s('/發招$ <%=rs("武功")%>')">
<input type=hidden name=wg1 value='<%=rs("武功")%>'>
<%=rs("武功")%></a></div>
</td>
<td height="20"><%=rs("級別")%></td>
<td height="20"><%=rs("內力")%></td>
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
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >

<tr>
<td colspan="3" height="17" bgcolor="#FF6633">
<div align="center"><font color="#FFFFFF">
<input type="checkbox" name="liji" value="on">
<font color="#000000">立即攻擊</font> </font></div>
</td>
</tr>
</table>
</div>
<p class=p1 align="center"><font color="#FFFFFF" size="2">殺傷力等於<br>
本人（等級*內力+武功+攻擊力）<br>
- <br>
對方（等級*內力+武功+防御力）</font></p>
</body>
</html>
