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

<title>怪物管理</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+6">[怪物管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="54"> 
      <div align="center"><font color="#FF6600">ID</font></div>
    </td>
    <td height="27" width="94"> 
      <div align="center"><font color="#FF6600">怪物名</font></div>
    </td>
    
    <td height="27" width="123"> 
      <div align="center"><font color="#FF6600">圖片說明</font></div>
    </td>
    <td height="27" width="71" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600">攻擊</font></div>
    </td>
    <td height="27" width="61"> 
      <div align="center"><font color="#FF6600">防御</font></div>
    </td>
    <td height="27" width="66"> 
      <div align="center"><font color="#FF6600">體力</font></div>
    </td>
    <td height="27" width="56"> 
      <div align="center"><font color="#FF6600">kill</font></div> 
</td>
   <td height="27" width="56"> 
      <div align="center"><font color="#FF6600">操作</font></div> 
</td>
     <td height="27" width="141"> 
      <div align="center"><font color="#FF6600">執行</font></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM 怪物區 ",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='ljgwuok.asp?a=m&id=<%=rs("id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="54"> 
        <div align="center"><font color="#FF6600">ID:</font> <font color="#FF6600"><%=rs("ID")%></font></div>
        <div align="center"></div>
      </td>
      <td height="2" width="94"> 
        <div align="center"> 
          <input type="text" name="a" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12"
value="<%=rs("怪物")%>" maxlength="255">
        </div>
      </td>
     
      <td height="2" width="123"> 
        <div align="center"> <img src="../ico/<%=rs("說明")%>.gif" border="0" alt="<%=rs("怪物")%> "> 
          <input type="text" name="b" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="8"
value="<%=rs("狀態")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="71"> 
        <div align="center"> 
          <input type="text" name="c" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("攻擊")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="61"> 
        <div align="center"> 
          <input type="text" name="d" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("防御")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="66"> 
        <div align="center"> 
          <input type="text" name="e" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("體力")%>" maxlength="255">
        </div>
      </td>
      <td height="2" width="56"> 
        <div align="center"> 
          <input type="text" name="f" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
value="<%=rs("kill")%>" maxlength="255">
        </div>
      </td>
<td height="2" width="56"> 
        <div align="center"> 
          <select name="select" style="border: 1px solid; background-color:#EEFFEE;font-size: 9pt; border-color:
#000000 solid" >
          <option value="aa" selected>乘以2</option>
          <option value="bb">乘以4</option>
          <option value="cc">乘以6</option>
          <option value="dd">乘以8</option>
          <option>------</option>
          <option value="ee">除以2</option>
          <option value="ff">除以4</option>
          <option value="gg">除以6</option>
	  <option value="hh">除以8</option>
        </select>
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
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<50 then
%>
  <form method=POST action='ljroomok.asp?a=n'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td width="54" height="2"> 
        <div align="center"><font color="#FF6600">添加寵物</font></div>
      </td>
      <td width="94" height="2"> 
        <div align="center"> 
          <input type="text" name="a" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="12"
maxlength="12">
        </div>
      </td>
      <td width="94" height="2"> 
        <div align="center"> 
          <select name="select">
            <option value="體力型" selected>體力型</option>
            <option value="內力型">內力型</option>
            <option value="防御型">防御型</option>
            <option value="武功型">武功型</option>
            <option value="法力型">法力型</option>
          </select>
        </div>
      </td>
      <td width="123" height="2"> 
        <div align="center"> 
          <input type="text" name="b" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
maxlength="255">
        </div>
      </td>
      <td width="71" height="2"> 
        <div align="center"> 
          <input type="text" name="c" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
maxlength="255">
        </div>
      </td>
      <td width="61" height="2"> 
        <div align="center"> 
          <input type="text" name="d" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
maxlength="255">
        </div>
      </td>
      <td width="66" height="2"> 
        <div align="center"> 
          <input type="text" name="e" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
maxlength="255">
        </div>
      </td>
      <td width="56" height="2"> 
        <div align="center"> 
          <input type="text" name="f" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="6"
maxlength="255">
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
</body>
</html>