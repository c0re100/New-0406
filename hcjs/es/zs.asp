<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();}</script>"
	Response.End 
end if

sl=abs(int(request.querystring("sl")))
id=lcase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
rs.open "select id,物品名 from 物品 where id=" & id & " and 類型<>'卡片' and 數量>="&sl&" and 擁有者="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：贈送不成功，你沒有這樣的物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wpname=rs("物品名")
%>
<title>江湖物品贈送</title>
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
<body bgcolor="#660000" background="../../chat/bk.jpg">
<table border=1 bgcolor="#BEE0FC" align=center width=336 cellpadding="10" cellspacing="13" height="293">
<tr valign="top">
    <td bgcolor=#99CCFF height="226"> 
      <div align="center">
<p><font size="3" color="#000000">物品贈送</font><font color="#FF3333" size="3"><br>
</font> </p>
<p><font color="#FF3333" size="3"> </font> </p>
</div>
<table width="324" height="212">
<tr>
<td height="179">
<form method=POST action='zs1.asp?id=<%=rs("id")%>&sl=<%=sl%>'>
<div align="left">
<p><font size="3" color="#000000">贈</font> 送 人<font color="#0000FF">:<%=name%>
</font><br>
<font size="3" color="#000000">贈</font>送物品<b><font color="#FF0000">:<%=wpname%><br>
</font></b>物品數量<b><font color="#FF0000">:<%=sl%></font></b>
個<br>
<font size="3" color="#000000">贈</font>送人:
                  <input type=text name=name size=14 maxlength="10" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" value="你朋友的名字">
<br>
<font size="3" color="#000000">贈</font> 言 :
<input type="text" name="zy" size="30" value="寫上你的贈言" style="color: #000000; background-color: #88B0D8; text-decoration: blink; border: 1px solid #000080" onMouseOver="this.style.backgroundColor = '#FFCC00'" onMouseOut="this.style.backgroundColor = '#88B0D8'" maxlength="30">
&nbsp; </p>
<p align="center">
<input type=submit value=確定贈送 name="submit" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
<input type=button value=返回上頁 onClick="javascript:history.go(-1)" name="button" style="border: 1px solid; font-size: 9pt; border-color:
#000000 solid">
</p>
</div>
</form>
</td>
</tr>
</table>
    </td>
</tr>
</table>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
%>