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
tiaojian=Request.Form("tiaojian")
show=trim(Request.Form("show"))
fs=int(Request.form("seekfs"))
%>
<html>
<head>
<title>�Τ��Ƭd�ݵ{��</title>
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body text="#000000" vlink="#FF9900" topmargin="0"
leftmargin="0" background="0064.jpg">
<p align="center"> <font color="#CC0000" face="����"><a href="javascript:this.location.reload()">��s</a></font> 
  <br>
  �o�@�ǬO�������󪺤H�I�I�m�W�i��ק�I<br>
  <font color="#FF0000"><b><%=tiaojian%></b></font> <br>
 <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
select case fs
case 0
	if show<>"" then
		rs.open "SELECT * FROM �Τ� where "& tiaojian &" order by -"&show,conn
	else
		rs.open "SELECT * FROM �Τ� where "& tiaojian &" order by lasttime",conn
	end if
%>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center">
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">�F�C����</font></b></font></td>
</tr>
<tr align="center">
<td>
      <table width="485" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr> 
          <td width="28" height="10"> 
            <div align="center"><font color="#FFFFFF">ID</font></div>
          </td>
          <td width="47" height="10"> 
            <div align="center"><font color="#FFFFFF">�m�W</font></div>
          </td>
          <td width="25" height="10"> 
            <div align="center"><font color="#FFFFFF">�ʧO</font></div>
          </td>
          <td width="63" height="10"> 
            <div align="center"><font color="#FFFFFF">����</font></div>
          </td>
          <td width="54" height="10"> 
            <div align="center"><font color="#FFFFFF">����</font></div>
          </td>
          <td width="82" height="10"> 
            <div align="center"><font color="#FFFFFF">�̫�n���ɶ�</font></div>
          </td>
          <td width="58" height="10"> 
            <div align="center"><font color="#FFFFFF">���򵥯�</font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="10"> 
            <div align="center"><font color="#FFFF00"><b><%=show%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="10"> 
            <div align="center"><font color="#FFFFFF">�n��</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="28" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
          </td>
          <td width="47" height="30"> 
            <div align="center"><a href="SHOWUSER.asp?username=<%=rs("�m�W")%>"><font color="#FF9900"><%=rs("�m�W")%></font></a></div>
          </td>
          <td width="25" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("�ʧO")%></font></div>
          </td>
          <td width="63" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
          </td>
          <td width="54" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
          </td>
          <td width="82" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("lasttime")%></font></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
          </td>
          <%if show<>"" then%>
          <td width="73" height="30"> 
            <div align="center"><font color="#FFFF00"><b><%=rs(show)%></b></font></div>
          </td>
          <%end if%>
          <td width="35" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("times")%></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
conn.close
%>
      </table>
</td>
</tr>
</table>
<div align="center"><font color="#000000">����H��:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">�H</font> 
</div>
<% case 1
sql="SELECT * FROM ���~ where "& tiaojian &" order by ����,�ƶq"
Set Rs=conn.Execute(sql)
%>
<table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
  <tr align="center"> 
    <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">�F�C����</font></b></font></td>
  </tr>
  <tr align="center"> 
    <td> 
      <table width="485" border="1" cellpadding="0" cellspacing="0" height="13">
        <tr> 
          <td width="26" height="7"> 
            <div align="center"><font color="#FFFFFF">ID</font></div>
          </td>
          <td width="41" height="7"> 
            <div align="center"><font color="#FFFFFF">���~�W</font></div>
          </td>
          <td width="52" height="7"> 
            <div align="center"><font color="#FFFFFF">�֦���</font></div>
          </td>
          <td width="32" height="7"> 
            <div align="center"><font color="#FFFFFF">����</font></div>
          </td>
          <td width="110" height="7"> 
            <div align="center"><font color="#FFFFFF">����</font></div>
          </td>
          <td width="34" height="7"> 
            <div align="center"><font color="#FFFFFF">�˳�</font></div>
          </td>
         
          <td width="53" height="7"> 
            <div align="center"><b><font color="#FFFFFF">�Ȩ�</font></b></div>
          </td>
          <td width="58" height="7"> 
            <div align="center"><font color="#FFFFFF">�ƶq</font></div>
          </td>
         
          <td width="59" height="7"> 
            <div align="center"><font color="#FFFFFF">�ާ@</font></div>
          </td>
        </tr>
        <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
        <tr> 
          <td width="26" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
          </td>
          <td width="41" height="30"> 
            <div align="center"><font color="#FF9900"><%=rs("���~�W")%></font></div>
          </td>
          <td width="52" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("�֦���")%></font></div>
          </td>
          <td width="32" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
          </td>
          <td width="110" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("����")%></font></div>
          </td>
          <td width="34" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("�˳�")%></font></div>
          </td>
          
          <td width="53" height="30"> 
            <div align="center"><b><font color="#FFFFFF"><%=rs("�Ȩ�")%></font></b></div>
          </td>
          <td width="58" height="30"> 
            <div align="center"><font color="#FFFFFF"><%=rs("�ƶq")%></font></div>
          </td>
         
          <td width="59" height="30"> 
            <div align="center"><font size="-1"><a href="../chat/del.asp?id=<%=rs("id")%>"><font color="#FFFF00">�R��</font></a></font></div>
          </td>
        </tr>
        <%
rs.movenext
loop
'conn.close
 case 2
sql="SELECT * FROM �׽m���~ where "& tiaojian &" order by �ƶq"
Set Rs=conn.Execute(sql)
%>
      </table>
    </td>
  </tr>
</table>
<div align="center"><font color="#000000">����H��:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">�H</font> 
  <br>
  <table border="0" width="500" cellspacing="0" cellpadding="0"
background="../jhmp/bg.gif" align="center">
    <tr align="center"> 
      <td background="../jhmp/top1.gif" width="500" height="26"><font
color="#FF6600"><b><font size="+1">�F�C����</font></b></font></td>
    </tr>
    <tr align="center"> 
      <td> 
        <table width="485" border="1" cellpadding="0" cellspacing="0" height="13">
          <tr> 
            <td width="32" height="12"> 
              <div align="center"><font color="#FFFFFF">ID</font></div>
            </td>
            <td width="64" height="12"> 
              <div align="center"><font color="#FFFFFF">���~�W</font></div>
            </td>
            <td width="71" height="12"> 
              <div align="center"><font color="#FFFFFF">�֦���</font></div>
            </td>
            <td width="61" height="12"> 
              <div align="center"><font color="#FFFFFF">�ƶq</font></div>
            </td>
            <td width="133" height="12"> 
              <div align="center"><font color="#FFFFFF">�\��</font></div>
            </td>
            <td width="55" height="12"> 
              <div align="center"><font color="#FFFFFF">�W�[�ƭ�</font></div>
            </td>
            <%if show<>"" then%>
            <%end if%>
            <td width="53" height="12"> 
              <div align="center"><font color="#FFFFFF">�ާ@</font></div>
            </td>
          </tr>
          <%
jl=0
do while not rs.bof and not rs.eof
jl=jl+1
%>
          <tr> 
            <td width="32" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("ID")%></font></div>
            </td>
            <td width="64" height="16"> 
              <div align="center"><font color="#FF9900"><%=rs("���~�W")%></font></div>
            </td>
            <td width="71" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("�֦���")%></font></div>
            </td>
            <td width="61" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("�ƶq")%></font></div>
            </td>
            <td width="133" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("�\��")%></font></div>
            </td>
            <td width="55" height="16"> 
              <div align="center"><font color="#FFFFFF"><%=rs("�W�[�ƭ�")%></font></div>
            </td>
            <%if show<>"" then%>
            <%end if%>
            <td width="53" height="16"> 
              <div align="center"><font size="-1"><a href="../chat/del1.asp?id=<%=rs("id")%>"><font color="#FFFF00">�R��</font></a></font></div>
            </td>
          </tr>
          <%
rs.movenext
loop
conn.close
%>
        </table>
      </td>
    </tr>
  </table>
  <div align="center"><font color="#000000">����H��:</font><b><font color="#0000FF"><%=(jl)%></font></b><font color="#000000">�H</font> 
  </div>
  <br>
</div>
<%end select%>
</body>
</html> 