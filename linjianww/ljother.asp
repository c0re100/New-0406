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

<title>�����޲z</title>
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
      <div align="center"><b><font color="#FF0000" size="+6">[�����޲z]</font></b></div>
    </td>
  </tr>
  <%Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select * from ���X where id>=0 order by id",conn
do while not rs.eof and not rs.bof
%>
  <form method=POST action='ljotherok.asp'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2">���_�}�� 
        <input type="text" name="wa" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
value="<%=rs("wabao")%>" maxlength="2">
        1--------- ���_�ȹ�|���}�� 0-----------���_��Ҧ��H�}��</td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> �|����d�� 
          <input type="text" name="hy" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
value="<%=rs("hysee")%>" maxlength="2">
          1------- �|���C���[�ݶȹ�|���}�� 0---------�|���C���[�ݹ�Ҧ��H�}��</p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <p> ���_�}��
<input type="text" name="lianbao" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" size="2"
value="<%=rs("lianbao")%>" maxlength="2">
          1------- ���_�ȹ�|���}�� 0---------���_��Ҧ��H�}��</p>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2"> 
        <div align="center"> 
          <input type="submit" value="�ק��s"
name="submit">
        </div>
      </td>
    </tr>
    <%
  rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
  </form>
</table>
<p class="p1" align="center">�H�W�Ů�B���ର��<br>
</p>
</body>
</html>