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
nickname=info(0)
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("../wb/ljjmwb.asa")
Conn.Open connstr
sql="select ���a�I�� from mess "
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Redirect "modifywb.asp"
Response.End
else
%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="setup.css">
</head>

<body text="#0033FF" link="#0000FF" vlink="#0000FF" alink="#0000FF" background="0064.jpg">
<div align="center"><font color="#FF0000" size="+6">���a�p����Ƨ��</font></div>
<form method="post" action="modifywbok.asp">
  <table border="1" cellspacing="0" align="center" cellpadding="0" width="560" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
    <tr bgcolor="#99CCFF"> 
      <td colspan="2"> 
        <p><font size="2" color="#FF0000">�`�N�G�ƾڮw��s�{�Ǧ]���ɶ������A�S���[�J���ŭȮɪ��˴��A�Фj�a���ɪ`�N�S���Ȫ��a���0<br>
          �ݭn���^������J�r��&lt;br&gt;<br>
          </font> </p>
      </td>
    </tr>
    <td width="84"><font size="-1" color="#FF6600">���a�I��</font></td>
    <td width="393"><font color="#FFCC99" size="-1"> 
      <input type="text" name="wbfk" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" value="<%=rs("���a�I��")%>" size="60">
      </font></td>
 
    
    <tr bgcolor="#FF6600"> 
      <td colspan="2"> 
        <div align="center"> 
          <input type="submit" value="�T �w">
          <font color="#CCCCCC">------- </font> 
          <input  onClick="javascript:window.document.location.href='modifywb.asp'" value="�� �s" type=button name="back">
        </div>
      </td>
    </tr>
  </table>
</form>
<%
end if
rs.close
%>
<table border="1" align="center" width="600" cellpadding="0"
cellspacing="0">
  <tr bgcolor="#99CCFF"> 
    <td height="17" colspan="6" bgcolor="#FF6600"> 
      <p align="center"><font color="#FFFFFF">�{���[�����a</font></p>
    
</table>
<table border="1" align="center" width="600" cellpadding="0"
cellspacing="0">
  <tr bgcolor="#C4DEFF"> 
    <td height="18" width="51"> 
      <div align="center"> <font size="2" color="#000000"><u>�s��</u></font></div>
    </td>
    <td height="18" width="102"> 
      <div align="center"><font size="2" color="#000000">���a�W��</font></div>
    </td>
    <td height="18" width="201"> 
      <div align="center"><font size="2" color="#000000">���a�a�}</font></div>
    </td>
    <td height="18" width="35"> 
      <div align="center"><font size="2" color="#000000">�H��</font></div>
    </td>
    <td height="18" width="95"> 
      <div align="center"><font size="2" color="#000000">�[�J���</font></div>
    </td>
    <td height="18" width="40"> 
      <div align="center"> <font size="2" color="#000000">�|��</font> </div>
    </td>
    <td height="18" width="60"> 
      <div align="center"> <font size="2" color="#000000">�� �@</font> </div>
    </td>
  </tr>
  <%
sql="select * from bar "
Set Rs=conn.Execute(sql)
do while not rs.bof and not rs.eof
%>
  <tr bgcolor="#FFFFFF"  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td height="11" width="51"><font size="2"><%=rs("id")%></font> <font size="2"> 
      <%if rs("qdlm")=true then%>
      <img src="../wb/hy.gif" width="16" height="18">
      <%end if%>
      </font> 
    <td height="11" width="102"> 
      <div align="center"><font color="#FFCC99" size="-1"> </font><font size="2"><%=rs("barname")%></font><font color="#FFCC99" size="-1"> 
        </font></div>
    </td>
    <td height="11" width="201"><font color="#FFCC99" size="-1"> </font><font size="2"><%=rs("add")%></font></td>
    <td height="11" width="35"><font color="#FFCC99" size="-1"> </font><font size="2"><%=rs("count")%></font><font color="#FFCC99" size="-1"> 
      </font><font size="2"> </font></td>
    <td height="11" width="95"><font color="#FFCC99" size="-1"> </font><font size="2"><%=rs("date")%></font></td>
    <td height="11" width="40"> <font size="2"> <%=rs("qdlm")%> </font></td>
    <td height="11" width="60"> 
      <div align="center"><font size="2"><a href="modifywbx.asp?wbid=<%=rs("id")%>">�J��</a> 
        <a href="modifywbdel.asp?id=<%=rs("id")%>">�R��</a></font></div>
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
</table>
<div align="center"></div>
