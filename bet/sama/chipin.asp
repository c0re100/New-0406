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
sqlstr="select 뽁ⓥ from Ξㅱ where id="&info(9)
rs.open sqlstr,conn
yin=rs("뽁ⓥ")
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
<td width="27">0많</td>
<td width="34">
<input  maxlength=4 size=4  name="horse0" >
</td>
<td width="24">1많 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse1" >
</td>
<td width="24">2많 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse2" >
</td>
<td width="24">3많 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse3" >
</td>
<td width="24">4많</td>
<td width="34">
<input  maxlength=4 size=4  name="horse4" >
</td>
<td width="24">5많 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse5" >
</td>
<td width="24">6많 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse6" >
</td>
<td width="24">7많</td>
<td width="34">
<input  maxlength=4 size=4  name="horse7" >
</td>
<td width="24">8많 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse8" >
</td>
<td width="24">9많</td>
<td width="34">
<input  maxlength=4 size=4  name="horse9" >
</td>
<td width="32">10많</td>
<td width="34">
<input  maxlength=4 size=4  name="horse10" >
</td>
<td width="32">11많 </td>
<td width="7">
<input  maxlength=4 size=4  name="horse11" >
</td>
</tr>
<tr>
<td colspan=10 height="16">
<div align="center"><font size="-1">힓퍬┸댔좼감놓좮</div>
</td>
<td colspan=4 align=right valign="top" height="16">
<div align="center"><font size="-1">
<input type=submit value="짾 �`" name="submit">
<input type=button onClick=javascript:top.window.close(); value='츙 낚' name="button">
</font></div>
</td>
<td colspan=10 height="16">
<div align="center"><span style='font-size:9pt'>Ξㅱ좬<%=info(0)%>
  뽁ⓥ좬<%=yin%></span></div>
</td>
</tr>
</table>
</FORM>
</body>
