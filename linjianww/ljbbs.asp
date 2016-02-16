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

<title>江湖論壇在線管理</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+3">[江湖論壇在線管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="670"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="85"> 
      <div align="center"><font color="#FF6600"> 論壇名</font></div>
    </td>
    <td height="27" width="160"> 
      <div align="center"><font color="#FF6600"> 論壇說明</font></div>
    </td>
    <td height="27" width="154" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600"> 版主</font></div>
    </td>
    <td height="27" width="141"> 
      <div align="center"><font color="#FF6600"> 站 長 操 作 </font></div>
    </td>
  </tr>
  <%Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../bbs/bbs.asa")
Conn.Open connstr
sql="select * from board where bid>=1 order by bid"
set rs=Conn.execute(sql)
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='ljbbsok.asp?a=m&id=<%=rs("bid")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="85"> 
        <div align="center"> 
          <input type="text" name="leixin" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("cname")%>" maxlength="240">
        </div>
      </td>
      <td height="2" width="160"> 
        <div align="center"> 
          <input type="text" name="name" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("bdinfo")%>" maxlength="240">
        </div>
      </td>
      <td height="2" width="154"> 
        <div align="center"> 
          <input type="text" name="wz1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("版主")%>" maxlength="240">
        </div>
      </td>
      <td height="2" width="141"> 
        <div align="center"> 
          <input type="submit" value="修改"
name="submit">
          &nbsp;&nbsp; 
          <input type="submit" value="刪除"
name="submit" >
        </div>
      </td>
    </tr>
  </form>
  <%
rs.movenext
loop
if s<50 then
%>
  <form method=POST action='ljbbsok.asp?a=n'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="85"> 
        <div align="center"> 
          <input type="text" name="leixin" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="240">
        </div>
      </td>
      <td height="2" width="160"> 
        <div align="center"> 
          <input type="text" name="name" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
maxlength="240">
        </div>
      </td>
      <td width="154" height="2"> 
        <div align="center"> 
          <input type="text" name="wz1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
maxlength="240">
        </div>
      </td>
      <td width="141" height="2"> 
        <div align="center"> 
          <input type="submit" value="添加"
name="submit">
        </div>
      </td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center"><font color="#009900"><b><font color="#FF0000" size="+3">[禁言管理]</font></b></font><br>
<%
rs.close
sql="select name from user_info"
set rs=Conn.execute(sql)
jingyan=rs("name")
set rs=nothing
conn.close
set conn=nothing%>
<form method=POST action='ljbbsok.asp?a=p'>
  <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="85"> 
        <div align="center"> 
          
        <input type="text" name="jingyan" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="30"
maxlength="240" value="<%=jingyan%>">
      </div>
      </td>
      
    <td height="2" width="160">&nbsp;</td>
      <td width="154" height="2"> 
        
      <div align="center"> </div>
      </td>
      <td width="141" height="2"> 
        <div align="center"> 
          
        <input type="submit" value="更 新"
name="submit">
        </div>
      </td>
    </tr>
  </form>

<p class="p1" align="center">多個版主設置在名字中間加| 例：江湖總站|熊熊 以上空格處不能為空<br>
</body>
</html>