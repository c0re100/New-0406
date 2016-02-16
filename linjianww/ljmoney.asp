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

<title>工資管理</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<table cellpadding="1" cellspacing="0" border="1" align="center" width="670"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27"> 
      <div align="center"><b><font color="#FF0000" size="+6">[工資管理]</font><font color="#FF6600"><br>
        <br>
        基本工資+會員等級x2萬+弟子數x5千+介紹人數x1000+戰鬥等級x1500 </font></b></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select * from 集合 where id>=0 order by id",conn
do while not rs.eof and not rs.bof
%>
  <form method=POST action='ljmoneyok.asp'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2">官府基本工資 
        <input type="text" name="gf" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("gfmoney")%>" maxlength="50">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 掌門基本工資 
          <input type="text" name="zm" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("zmmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 長老基本工資 
          <input type="text" name="zl" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("zlmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 堂主基本工資 
          <input type="text" name="tz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("tzmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 弟子基本工資 
          <input type="text" name="dz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("dzmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 萬年靈芝領取個數 
          <input type="text" name="lz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("lznum")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> 魔力鑽石領取個數 
          <input type="text" name="zs" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("zsnum")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <div align="center"> 
          <input type="submit" value="修改更新"
name="submit">
        </div>
      </td>
    </tr> <%
  rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  </form>
 
</table>
<p class="p1" align="center">以上空格處不能為空</p>
</body>

</html>