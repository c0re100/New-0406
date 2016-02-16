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
rs.open "SELECT 身份,門派 FROM 用戶 where id="&info(9),conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
else
if rs("身份")<>"掌門" then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('不是掌門，不要亂闖！');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
pai=rs("門派")
rs.close
%>
<html>
<head>
<title><%=pai%>武功設置</title>
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
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: 新細明體; font-size: 9pt; font-style: bold;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<body background="../ff.jpg"><div align="center">
<table cellpadding="1" cellspacing="0" border="1" align="center" width="597"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#990000"> 
    <td width="64" height="27"> 
      <div align="center"><font color="#FFFFFF">ID號</font></div></td>
    <td width="84" height="27"> 
      <div align="center"><font color="#FFFFFF"> 招 式 
        名 </font></div></td>
    <td width="104" height="27"> 
      <div align="center"><font color="#FFFFFF"> 類 
        型</font></div></td>
    <td width="126" height="27"> 
      <div align="center"><font color="#FFFFFF"> 所 
        用 內 力 </font></div></td>
    <td width="197" height="27"> 
      <div align="center"><font color="#FFFFFF"> 掌 
        門 操 作 </font></div></td>
  </tr>
  <%
rs.open "SELECT id,武功,類型,內力 FROM 武功 where 門派='" & pai & "'",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='setwg1.asp?a=m&id=<%=rs("id")%>&pai=<%=pai%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="64"> <div align="center"><font color="#FF6600"><%=rs("id")%> </font> </div>
        <div align="center"></div></td>
      <td height="2" width="84"> <div align="center"> 
          <input type="text" name="wg1" size="10"
value="<%=rs("武功")%>" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="10">
        </div></td>
      <td height="2" width="104"> <div align="center"> 
          <input type="text" name="lx" size="4"
value="<%=rs("類型")%>" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="4">
        </div></td>
      <td height="2" width="126"> <div align="center"> 
          <input type="text" name="nl" size="8"
value="<%=rs("內力")%>" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="4">
        </div></td>
      <td height="2" width="197"> <div align="center"> 
          <input type="submit" value="修改"
name="submit">
          <input type="submit" value="刪除"
name="submit">
        </div></td>
    </tr>
  </form>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<10 then
%>
  <form method=POST action='setwg1.asp?a=n&wg=&pai=<%=pai%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td width="64" height="2"> <div align="center"><font color="#FF6600">新建招式</font></div></td>
      <td width="84" height="2"> <div align="center"> 
          <input type="text" name="wg1" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"
maxlength="10">
        </div></td>
      <td width="104" height="2"> <div align="center"> 
          <input type="text" name="lx" size="4" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"
maxlength="4">
        </div></td>
      <td width="126" height="2"> <div align="center"> 
          <input type="text" name="nl" size="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"
maxlength="4">
        </div></td>
      <td width="197" height="2"> <div align="center"> 
          <input type="submit" value="添加"
name="submit">
        </div></td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center">類型：攻擊、防御、恢復<br>
如果武功名字不文雅，官府將立刻刪除對應幫派，永不恢復！<br>
武功計算采用新的工式，現處測試中，如有什麼意見可以找官府提出！<br>
對於設置武功不要全用9999內力，可以根據實際情況進行選擇，最好為由小到大遞增！<br>
</body>
</html>
<%end if%>