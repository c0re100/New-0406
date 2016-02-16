<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<html>
<head>
<title><%=pai%>武功設置</title>
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
    <font color="#009900"><b><font color="#FF0000" size="+6">[ <b>武 功 管 理</b> 
    ] </font></b></font><font
color="#FFFFFF"><br>
    </font> </p>
  </div>
<table cellpadding="1" cellspacing="0" border="1" align="center" width="700"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27" width="69"> 
      <div align="center"><font color="#FF6600">ID號</font></div>
    </td>
    <td height="27" width="73"> 
      <div align="center"><font color="#FF6600"> 門 派</font></div>
    </td>
    <td height="27" width="94"> 
      <div align="center"><font color="#FF6600"> 招 式</font></div>
    </td>
    <td height="27" width="94"> 
      <div align="center"><font color="#FF6600"> 類型</font></div>
    </td>
    <td height="27" width="98"> 
      <div align="center"><font color="#FF6600"> 級別</font></div>
    </td>
    <td height="27" width="87"> 
      <div align="center"><font color="#FF6600"> 修煉級</font></div>
    </td>
    <td height="27" width="72"> 
      <div align="center"><font color="#FF6600"> 內 力</font></div>
    </td>
    <td height="27" width="79"> 
      <div align="center"><font color="#FF6600"> 操 作 </font></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM 武功 where 內力>1 order by id",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='ljsetwg1.asp?a=m&id=<%=rs("id")%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="69"> 
        <div align="center"><font color="#FF6600"><%=rs("id")%> </font> </div>
        <div align="center"></div>
      </td>
      <td height="2" width="73"> 
        <div align="center"> <font color="#FF6600"><%=rs("門派")%></font></div>
      </td>
      <td height="2" width="94"> 
        <div align="center"> 
          <input type="text" name="wg1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("武功")%>" maxlength="10">
        </div>
      </td>
      <td height="2" width="94"> 
        <div align="center"> 
          <input type="text" name="wg2" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("類型")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="98"> 
        <div align="center"> 
          <input type="text" name="wg3" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("級別")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="87"> 
        <div align="center"> 
          <input type="text" name="wg4" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("修煉級")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="72"> 
        <div align="center"> 
          <input type="text" name="nl" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
value="<%=rs("內力")%>" maxlength="4">
        </div>
      </td>
      <td height="2" width="79"> 
        <div align="center"> 
          <input type="submit" value="修改"
name="submit">
          <input type="submit" value="刪除"
name="submit" >
        </div>
      </td>
    </tr>
  </form>
  <%rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<50 then
%>
  <form method=POST action='ljsetwg1.asp?a=n'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td width="69" height="2"> 
        <div align="center"><font color="#FF6600">新建招式</font></div>
      </td>
      <td width="73" height="2"> 
        <div align="center"> 
          <input type="text" name="mp" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="10">
        </div>
      </td>
      <td width="94" height="2"> 
        <div align="center"> 
          <input type="text" name="wg1" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="10">
        </div>
      </td>
      <td width="94" height="2"> 
        <div align="center"> 
          <input type="text" name="wg2" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="4">
        </div>
      </td>
      <td width="98" height="2"> 
        <div align="center"> 
          <input type="text" name="wg3" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="4">
        </div>
      </td>
      <td width="87" height="2"> 
        <div align="center"> 
          <input type="text" name="wg4" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="4">
        </div>
      </td>
      <td width="72" height="2"> 
        <div align="center"> 
          <input type="text" name="nl" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="10"
maxlength="4">
        </div>
      </td>
      <td width="79" height="2"> 
        <div align="center"> 
          <input type="submit" value="添加"
name="submit">
        </div>
      </td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center">內力不可設置太多，內力太多修練到後來攻擊太強<br>
</body>

</html>