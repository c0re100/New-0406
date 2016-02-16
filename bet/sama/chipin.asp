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
sqlstr="select 銀兩 from 用戶 where id="&info(9)
rs.open sqlstr,conn
yin=rs("銀兩")
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
<td width="27">0號</td>
<td width="34">
<input  maxlength=4 size=4  name="horse0" >
</td>
<td width="24">1號 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse1" >
</td>
<td width="24">2號 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse2" >
</td>
<td width="24">3號 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse3" >
</td>
<td width="24">4號</td>
<td width="34">
<input  maxlength=4 size=4  name="horse4" >
</td>
<td width="24">5號 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse5" >
</td>
<td width="24">6號 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse6" >
</td>
<td width="24">7號</td>
<td width="34">
<input  maxlength=4 size=4  name="horse7" >
</td>
<td width="24">8號 </td>
<td width="34">
<input  maxlength=4 size=4  name="horse8" >
</td>
<td width="24">9號</td>
<td width="34">
<input  maxlength=4 size=4  name="horse9" >
</td>
<td width="32">10號</td>
<td width="34">
<input  maxlength=4 size=4  name="horse10" >
</td>
<td width="32">11號 </td>
<td width="7">
<input  maxlength=4 size=4  name="horse11" >
</td>
</tr>
<tr>
<td colspan=10 height="16">
<div align="center"><font size="-1">靈劍江湖賽馬場！</div>
</td>
<td colspan=4 align=right valign="top" height="16">
<div align="center"><font size="-1">
<input type=submit value="下 注" name="submit">
<input type=button onClick=javascript:top.window.close(); value='關 閉' name="button">
</font></div>
</td>
<td colspan=10 height="16">
<div align="center"><span style='font-size:9pt'>用戶：<%=info(0)%>
  銀兩：<%=yin%></span></div>
</td>
</tr>
</table>
</FORM>
</body>
