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
rs.open "select �|������,�|���O,���� from �Τ� where id="&info(9),conn
if rs("�|������")=0 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A���O�|���Ц^�a�I');location.href = '../help/info.asp?type=2&name=�|���[�J��k';}</script>"
response.end
end if
hy=rs("�|���O")
jb=rs("����")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�|�O�����</title>
<link REL="StyleSheet" HREF="../style.css" TITLE="Contemporary">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" LEFTMARGIN="0" TOPMARGIN="0" >
<div align="center"><center>
    <table border="0" width="513" cellspacing="0" cellpadding="0" ALIGN="CENTER">
      <tr> 
        <td align="center"> 
          <p><br>
            <FONT COLOR="#FF66FF" SIZE="2">�z�{���|�O�G </FONT><font color="#000000"><b><font color=#FFFFCC><%=hy%></font></b></font><FONT COLOR="#FF66FF" SIZE="2"> 
            �����G </FONT><font color="#000000"><b><font color=#FFFFCC><%=jb%> <br>
            �ഫ���v:1���|�O����100�Ӫ���</font></b></font></p>
          <form name="form1" method="post" action="hxiehyok.asp">
            <input type="text" name="text" value="�п�J�A�n�ഫ���|�O�I" maxlength="10" size="22">
            <input type="submit" name="Submit" value=" �� �� ">
          </form>
          <p>&nbsp;</p>
          </td>
      </tr>
      <tr> 
        <td align="center"></td>
      </tr>
    </table>
  </center>
</div>
</body>
</html>
<html>