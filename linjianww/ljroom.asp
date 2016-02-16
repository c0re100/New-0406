<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head>

<title>房間管理</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<div align="center"> 
  <p><br>
    <font color="#009900"><b><font color="#FF0000" size="+6">[房間管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="168"> 
      <div align="center"><font color="#FF6600">ID</font></div>
    </td>
    <td height="27" width="77"> 
      <div align="center"><font color="#FF6600">房間名</font></div>
    </td>
    <td height="27" width="114"> 
      <div align="center"><font color="#FF6600">人數</font></div>
    </td>
    <td height="27" width="63" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600">打鬥限制</font></div>
    </td>
    <td height="27" width="103"> 
      <div align="center"><font color="#FF6600">限制說明</font></div>
    </td>
    <td height="27" width="150"> 
      <div align="center"><font color="#FF6600">表達式</font></div>
    </td>
<td height="27" width="150"> 
      <div align="center"><font color="#FF6600">發招</font></div>
    </td>
<td height="27" width="150"> 
      <div align="center"><font color="#FF6600">事件</font></div>
    </td>
<td height="27" width="250"> 
      <div align="center"><font color="#FF6600">操作</font></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM room ",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='ljroomok.asp?a=m&id=<%=rs("id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="88"> 
        <div align="center"><font color="#FF6600">ID:</font> <font color="#FF6600"><%=rs("ID")%></font></div>
        <div align="center"></div>
      </td>
<td height="2" width="142"> 
        <div align="center"> 
          <input type="text" name="a" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12"
value="<%=rs("a")%>" maxlength="12">
        </div>
      </td>
     <td height="2" width="142"> 
        <div align="center"> 
          <input type="text" name="b" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4"
value="<%=rs("b")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="186"> 
        <div align="center"> 
          <input type="text" name="c" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
value="<%=rs("c")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="103"> 
        <div align="center"> 
          <input type="text" name="d" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("d")%>" maxlength="255">
        </div>
      </td>
<td height="2" width="103"> 
        <div align="center"> 
          <input type="text" name="e" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("e")%>" maxlength="255">
        </div>
      </td>
<td height="2" width="45"> 
        <div align="center"> 
          <input type="text" name="f" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4"
value="<%=rs("f")%>" maxlength="4">
        </div>
      </td>
<td height="2" width="4"> 
        <div align="center"> 
          <input type="text" name="g" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4"
value="<%=rs("g")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="150"> 
        <div align="center"> <a href="ljroomok.asp?a=m&id=<%=rs("id")%>"> 
          <input type="submit" value="修改"
name="submit">
          &nbsp;&nbsp; </a> 
          <input type="submit" value="刪除"
name="submit" >
        </div>
      </td>
    </tr>
  </form>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<50 then
%>
  <form method=POST action='ljroomok.asp?a=n'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td width="88" height="2"> 
        <div align="center"><font color="#FF6600">添加房間</font></div>
      </td>
      <td width="142" height="2"> 
<div align="center">
        <input type="text" name="a" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12"
maxlength="12"></div>
      </td>
<td width="142" height="2"> 
<div align="center"> 
        <input type="text" name="b" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4"
maxlength="10">
 </div>
      </td>
      <td width="186" height="2"> 
        <div align="center"> 
          <input type="text" name="c" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
maxlength="4">
        </div>
      </td>
      <td width="103" height="2"> 
        <div align="center"> 
          <input type="text" name="d" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="255">
        </div>
      </td>
 <td width="103" height="2"> 
        <div align="center"> 
          <input type="text" name="e" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="255">
        </div>
      </td>
 <td width="43" height="2"> 
        <div align="center"> 
          <input type="text" name="f" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4"
maxlength="4">
        </div>
      </td>
 <td width="4" height="2"> 
        <div align="center"> 
          <input type="text" name="g" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="4"
maxlength="4">
        </div>
      </td>
      <td width="150" height="2"> 
        <div align="center"> 
          <input type="submit" value="添加"
name="submit">
        </div>
      </td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center">注意：房間操作需要重置服務器纔能生效，即可以把根目錄下的global.asa 改成global.txt,過1、2分鐘再改回global.asa即上面的設置就能生效！<br>
</p>
</body>

</html>