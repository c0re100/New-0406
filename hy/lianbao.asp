<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "Select lianbao from ���X where id>=0 order by id",conn
lian=rs("lianbao")
rs.close
sql="select �|������,�_��,�k�O,���� from �Τ� where ID=" & info(9)
set rs=conn.execute(sql) 
zhizi=rs("����")
if rs("�|������")=0 and lian=1 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���O�|���Ц^�a�I');location.href = '../help/info.asp?type=2&name=�|���[�J��k';}</script>"
Response.End
end if
if zhizi<1000 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�����C��1000�L����O���_�I');window.close();}</script>"
Response.End
end if
if rs("�_��")="�L" then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���Ӫ��_���F�C�����ۧr�I');window.close();}</script>"
Response.End
end if

%>
<html>

<head>
<title>���_</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css"><!--td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { text-decoration: none; font-family: �s�ө���; font-size: 9pt }
a:hover      { text-decoration: underline; color: #CC0000; font-family: �s�ө���; font-size: 
               9pt }
.big         { font-family: �s�ө���; font-size: 12pt }
.txt         { font-family: �s�ө���; font-size: 10.8pt }
--></style>
</head>

<body bgcolor="#660000" LEFTMARGIN="0" TOPMARGIN="0">
<table border="1" bgcolor="#003300" align="center" width="389" cellpadding="8"
cellspacing="8">
  <tr bgcolor="#FFFFFF"> 
    <td height="99" bgcolor="#99CCFF"> 
      <table border="1"
      width="355" cellspacing="2" cellpadding="1" bordercolor="#65251C">
        <tr> 
          <td colspan="5" height="33"> 
            <div align="center"> <b><font size="+2" color="#FF0000">���_</font></b></div>
          </td>
        </tr>
        <tr> 
          <td width="113" height="17" bgcolor="#FFFFFF"><font size="2" class="c">�_���G</font></td>
          <td height="17" colspan="4" bgcolor="#FFFFFF"><%=rs("�_��")%> </td>
        </tr>
        <tr> 
          <td width="113" bgcolor="#FFFFFF" height="10"><font size="2">��e�k�O</font><font size="2" class="c">:</font></td>
          <td width="51" bgcolor="#FFFFFF" height="10"><%=rs("�k�O")%> </td>
          <td bgcolor="#FFFFFF" height="10"><font size="2">���G<%=rs("����")%></font></td>
        </tr>
      </table>
      <table width="355" border="1" align="center" cellspacing="1" cellpadding="1" bgcolor="#000000" bordercolor="#CCCCCC">
        <tr> 
          <td width="121"> 
            <div align="center"><a href="liangong.asp"><font color="#FFFFFF">�E<br>
              ��<br>
              �|<br>
              ��</font></a></div>
          </td>
          <td width="89"><a href="lianbaook.asp"><img src="111.gif" width="79" height="75" border="0" alt="�I�����_"></a></td>
          <td width="121"> 
            <div align="center"><font color="#FFFFFF">��<br>
              �L<br>
              ��<br>
              ��</font></div>
          </td>
        </tr>
      </table>
      <br>
      �ثe����|���}��A�A���F�_���]�F�C�����ۡ^��i�H�b�����_�l���_�����k�O,��o���k�O�h�֮ھڧA�����өw�I<br>
    
  </table>

</body>

</html>
<%
conn.close
set conn=nothing
set rs=nothing%>
