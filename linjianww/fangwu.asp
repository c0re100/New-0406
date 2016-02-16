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

<title>愛情小屋管理</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+3">[愛情小屋管理]</font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="600"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="85"> 
      <div align="center"><font color="#FF6600"> 類 型</font></div>
    </td>
    <td height="27" width="80"> 
      <div align="center"><font color="#FF6600"> 地 區 </font></div>
    </td>
    <td height="27" width="84" bgcolor="#99CCFF"> 
      <div align="center"><font color="#FF6600"> 售 價 </font></div>
    </td>
    <td height="27" width="88"> 
      <div align="center"><font color="#FF6600"> 增加房子數 </font></div>
    </td>
    <td height="27" width="81"> 
      <div align="center"><font color="#FF6600"> 站 長 操 作 </font></div>
    </td>
  </tr>
  
  <%

if s<50 then
%>
  <form method=POST action='fangwuok.asp?a=n'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="85"> 
        <div align="center"> 
         <select name="leixin" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="一般房屋" selected>一般房屋</option>
      <option value="高級公寓">高級公寓</option>
      <option value="花園洋房">花園洋房</option>
      <option value="豪華別墅">豪華別墅</option>
    </select>
        </div>
      </td>
      <td height="2" width="160"> 
        <div align="center"> 
          <select name="jiedao" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="平民區" selected>平民區</option>
      <option value="華京區">華京區</option>
      <option value="華翔區">華翔區</option>
      <option value="華貴區">華貴區</option>
    </select>
        </div>
      </td>
<td height="2" width="160"> 
        <div align="center"> 
          <select name="money" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="500000" selected>500000</option>
      <option value="1000000">1000000</option>
      <option value="3000000">3000000</option>
      <option value="5000000">5000000</option>
    </select>
        </div>
      </td>
      <td width="154" height="2"> 
        <div align="center"> 
         <select name="sn" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
      <option value="1" selected>1</option>
      <option value="5">5</option>
      <option value="10">10</option>
      <option value="15">15</option>
      <option value="20">20</option>
    </select>
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
<p class="p1" align="center">以上空格處不能為空<br>
</body>
</html>