<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sqlstr="select �Ȩ� from �Τ� where id="&info(9)
rs.open sqlstr,conn
yin=rs("�Ȩ�")
rs.Close
set rs=nothing
conn.close
set conn=nothing
%>
<head><title></title><LINK rel=stylesheet href='css.css'></head>
<bgsound src="horse.mid" loop="-1">
<body oncontextmenu=self.event.returnValue=false bgcolor="#CCCCCC" topmargin="5" >
<form method="post" action="compete.asp"  target=compfrm>
<table  border="0" cellspacing="0" cellpadding="0" bgcolor="#CCCCCC" width='95%' align=center>
<tr>
<td width="27">0��</td>
<td width="34">
<input  maxlength=4 size=4  name="horse0" >
</td>
<td width="24">1�� </td>
<td width="34">
<input  maxlength=4 size=4  name="horse1" >
</td>
<td width="24">2�� </td>
<td width="34">
<input  maxlength=4 size=4  name="horse2" >
</td>
<td width="24">3�� </td>
<td width="34">
<input  maxlength=4 size=4  name="horse3" >
</td>
<td width="24">4��</td>
<td width="34">
<input  maxlength=4 size=4  name="horse4" >
</td>
<td width="24">5�� </td>
<td width="34">
<input  maxlength=4 size=4  name="horse5" >
</td>
<td width="24">6�� </td>
<td width="34">
<input  maxlength=4 size=4  name="horse6" >
</td>
<td width="24">7��</td>
<td width="34">
<input  maxlength=4 size=4  name="horse7" >
</td>
<td width="24">8�� </td>
<td width="34">
<input  maxlength=4 size=4  name="horse8" >
</td>
<td width="24">9��</td>
<td width="34">
<input  maxlength=4 size=4  name="horse9" >
</td>
<td width="32">10��</td>
<td width="34">
<input  maxlength=4 size=4  name="horse10" >
</td>
<td width="32">11�� </td>
<td width="7">
<input  maxlength=4 size=4  name="horse11" >
</td>
</tr>
<tr>
<td colspan=10 height="16">
<div align="center"><font size="-1">�F�C�����ɰ����I</div>
</td>
<td colspan=4 align=right valign="top" height="16">
<div align="center"><font size="-1">
<input type=submit value="�U �`" name="submit">
<input type=button onClick=javascript:top.window.close(); value='�� ��' name="button">
</font></div>
</td>
<td colspan=10 height="16">
<div align="center"><span style='font-size:9pt'>�Τ�G<%=info(0)%>
  �Ȩ�G<%=yin%></span></div>
</td>
</tr>
</table>
</FORM>
</body>
