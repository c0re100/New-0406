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

<title>�u��޲z</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body background="0064.jpg">
<table cellpadding="1" cellspacing="0" border="1" align="center" width="670"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#99CCFF"> 
    <td height="27"> 
      <div align="center"><b><font color="#FF0000" size="+6">[�u��޲z]</font><font color="#FF6600"><br>
        <br>
        �򥻤u��+�|������x2�U+�̤l��x5�d+���ФH��x1000+�԰�����x1500 </font></b></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select * from ���X where id>=0 order by id",conn
do while not rs.eof and not rs.bof
%>
  <form method=POST action='ljmoneyok.asp'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2">�x���򥻤u�� 
        <input type="text" name="gf" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("gfmoney")%>" maxlength="50">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> �x���򥻤u�� 
          <input type="text" name="zm" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("zmmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> ���Ѱ򥻤u�� 
          <input type="text" name="zl" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("zlmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> ��D�򥻤u�� 
          <input type="text" name="tz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("tzmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> �̤l�򥻤u�� 
          <input type="text" name="dz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("dzmoney")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> �U�~�F�ۻ���Ӽ� 
          <input type="text" name="lz" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("lznum")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> �]�O�p�ۻ���Ӽ� 
          <input type="text" name="zs" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="20"
value="<%=rs("zsnum")%>" maxlength="50">
        </p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <div align="center"> 
          <input type="submit" value="�ק��s"
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
<p class="p1" align="center">�H�W�Ů�B���ର��</p>
</body>

</html>