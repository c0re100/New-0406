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
rs.open "SELECT ����,���� FROM �Τ� where id="&info(9),conn
if rs.bof or rs.eof then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Redirect "error.asp?id=453"
	response.end
else
if rs("����")<>"�x��" then
	rs.close
	conn.close
	set conn=nothing
	set rs=nothing
	Response.Write "<script language=JavaScript>{alert('���O�x���A���n�����I');location.href = 'javascript:history.back()';}</script>"
	Response.End
end if
pai=rs("����")
rs.close
%>
<html>
<head>
<title><%=pai%>�Z�\�]�m</title>
<style>
A:visited{TEXT-DECORATION: none ;color:005FA2}
A:active{TEXT-DECORATION: none;color:005FA2}
A:link{text-decoration: none;color:005FA2}
A:hover { color: #C81C00; POSITION: relative; BORDER-BOTTOM: #808080 1px dotted A:hover ;}
.t{LINE-HEIGHT: 1.4}
BODY{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt;color:009FE0
scrollbar-face-color:#6080B0;scrollbar-shadow-color:#FFFFFF;scrollbar-highlight-color:#FFFFFF;
scrollbar-3dlight-color:#FFFFFF;scrollbar-darkshadow-color:#FFFFFF;scrollbar-track-color:#FFFFFF;
scrollbar-arrow-color:#FFFFFF;
SCROLLBAR-HIGHLIGHT-COLOR: buttonface;SCROLLBAR-SHADOW-COLOR: buttonface;
SCROLLBAR-3DLIGHT-COLOR: buttonhighlight;SCROLLBAR-TRACK-COLOR: #eeeeee;
SCROLLBAR-DARKSHADOW-COLOR: buttonshadow}
TD,DIV,form ,OPTION,P,TD,BR{FONT-FAMILY: �s�ө���; FONT-SIZE: 9pt}
INPUT{BORDER-TOP-WIDTH: 1px; PADDING-RIGHT: 1px; PADDING-LEFT: 1px;BORDER-LEFT-WIDTH: 1px; FONT-SIZE: 9pt; BORDER-LEFT-COLOR: #000000; BORDER-BOTTOM-WIDTH: 1px; BORDER-BOTTOM-COLOR: #000000; PADDING-BOTTOM: 1px; BORDER-TOP-COLOR: #000000; PADDING-TOP: 1px; HEIGHT: 18px; BORDER-RIGHT-WIDTH: 1px; BORDER-RIGHT-COLOR: #000000}
textarea, select {border-width: 1; border-color: #000000; background-color: #efefef; font-family: �s�ө���; font-size: 9pt; font-style: bold;}
</style>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<body background="../ff.jpg"><div align="center">
<table cellpadding="1" cellspacing="0" border="1" align="center" width="597"
bordercolorlight="#000000" bordercolordark="#FFFFFF">
  <tr bgcolor="#990000"> 
    <td width="64" height="27"> 
      <div align="center"><font color="#FFFFFF">ID��</font></div></td>
    <td width="84" height="27"> 
      <div align="center"><font color="#FFFFFF"> �� �� 
        �W </font></div></td>
    <td width="104" height="27"> 
      <div align="center"><font color="#FFFFFF"> �� 
        ��</font></div></td>
    <td width="126" height="27"> 
      <div align="center"><font color="#FFFFFF"> �� 
        �� �� �O </font></div></td>
    <td width="197" height="27"> 
      <div align="center"><font color="#FFFFFF"> �x 
        �� �� �@ </font></div></td>
  </tr>
  <%
rs.open "SELECT id,�Z�\,����,���O FROM �Z�\ where ����='" & pai & "'",conn
s=0
do while not rs.eof and not rs.bof
s=s+1
%>
  <form method=POST action='setwg1.asp?a=m&id=<%=rs("id")%>&pai=<%=pai%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td height="2" width="64"> <div align="center"><font color="#FF6600"><%=rs("id")%> </font> </div>
        <div align="center"></div></td>
      <td height="2" width="84"> <div align="center"> 
          <input type="text" name="wg1" size="10"
value="<%=rs("�Z�\")%>" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="10">
        </div></td>
      <td height="2" width="104"> <div align="center"> 
          <input type="text" name="lx" size="4"
value="<%=rs("����")%>" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="4">
        </div></td>
      <td height="2" width="126"> <div align="center"> 
          <input type="text" name="nl" size="8"
value="<%=rs("���O")%>" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid" maxlength="4">
        </div></td>
      <td height="2" width="197"> <div align="center"> 
          <input type="submit" value="�ק�"
name="submit">
          <input type="submit" value="�R��"
name="submit">
        </div></td>
    </tr>
  </form>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
if s<10 then
%>
  <form method=POST action='setwg1.asp?a=n&wg=&pai=<%=pai%>'>
    <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
      <td width="64" height="2"> <div align="center"><font color="#FF6600">�s�ةۦ�</font></div></td>
      <td width="84" height="2"> <div align="center"> 
          <input type="text" name="wg1" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"
maxlength="10">
        </div></td>
      <td width="104" height="2"> <div align="center"> 
          <input type="text" name="lx" size="4" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"
maxlength="4">
        </div></td>
      <td width="126" height="2"> <div align="center"> 
          <input type="text" name="nl" size="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid"
maxlength="4">
        </div></td>
      <td width="197" height="2"> <div align="center"> 
          <input type="submit" value="�K�["
name="submit">
        </div></td>
    </tr>
  </form>
  <%end if%>
</table>
<p class="p1" align="center">�����G�����B���s�B��_<br>
�p�G�Z�\�W�r���嶮�A�x���N�ߨ�R�����������A�ä���_�I<br>
�Z�\�p����ηs���u���A�{�B���դ��A�p������N���i�H��x�����X�I<br>
���]�m�Z�\���n����9999���O�A�i�H�ھڹ�ڱ��p�i���ܡA�̦n���Ѥp��j���W�I<br>
</body>
</html>
<%end if%>