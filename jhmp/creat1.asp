<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(3)<30 then Response.Redirect "../error.asp?id=460"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �x�� from ���� where �x��='"&info(0)&"'",conn
if not rs.bof or not rs.eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�w�g�O�x���F�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
rs.open "select ����,����,�Ȩ�,����,�|������ from �Τ� where id="&info(9),conn
if rs("����")="�x��" or rs("����")="�C�L" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�n�n���A�O�o�áI');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("�|������")<=1 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('���ŷ|���H�W�~�i�H�ӳЬ��I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("����")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x������԰����Ťj��100�šI');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("�Ȩ�")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x������Ȩ�j��2���I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
if rs("����")<10000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�x��������j��10000�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<html>
<head>
<title>�۳Ъ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>

<body bgcolor="#000000" text="#000000" link="#000080" alink="#800000" vlink="#000080" background="../linjianww/0064.jpg">
<form action="newmp.asp" method=POST >
  <ul>
    <table border=0 cellspacing=0 cellpadding=0 align="center" width="560" height="104">
      <tr> 
        <td height="49"> 
          <p align="center"><font size="2" class=c><br>
            </font><b><font size="+6" color="#FF0000">�۳Ъ���</font></b><font size="2" class=c><b><br>
            </b> <br>
            </font></p>
        </td>
      </tr>
      <tr> 
        <td height="2"> 
          <div align="center"><font color="#FFFFFF"><font color="#000000" size="+1">�����W:</font></font> 
            <input name="subject" size=27 style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength=10>
            <br>
            <br>
            �`�G�����W�ٳ̦h��10�ӭ^��w�šA5�Ӥ���r��<br>
            <br>
          </div>
        </td>
      </tr>
      <tr> 
        <td height="25"> 
          <div align="center"></div>
          <div align="center"><font color="#FFFFFF"><font color="#000000" size="+1">² 
            ��:</font></font> 
            <input name="comment" value="" size=27 style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength=30>
            <br>
            <br>
            �`�G����²���̦h��30�Ӧr�šA�B���i���šI<br>
          </div>
        </td>
      </tr>
      <tr> 
        <td height="25"> 
          <div align="center"> 
            <p><font color="#FFFFFF"><font color="#000000" size="+1"><br>
              �̤l�ʧO:</font></font> 
              <select name="ljxb" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
                <option value="�k">�k</option>
                <option value="�k">�k</option>
                <option value="�Ҧ�" selected>�Ҧ��H</option>
              </select>
              <br>
              <br>
              �`�G�̤l�ʧO���Ӫ��������\�k/�k/�Ҧ��H���[�J�I</p>
            <p>&nbsp;</p>
          </div>
        </td>
      </tr>
    </table>
    <div align="center"> <font size="3" class="c" color="#000000"><br>
      <br>
      <input type="HIDDEN" name="action" value="RegSubmit">
      <input type="SUBMIT" name="Submit" value="�� ��">
      <input type="RESET" name="Reset" value="�M ��">
      </font> </div>
  </ul>
</form>
</body>
</html>
