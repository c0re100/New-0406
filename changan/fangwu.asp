<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
name=info(0)%>
<!--#include file="data1.asp"--><%
sql="SELECT * FROM �Ы� where  ����='�@��Ы�'"
Set Rs=conn.Execute(sql)
%>
<html>
<link href="../chat/ccss2.css" rel="stylesheet" type="text/css">

<body bgcolor="#FFFDDF" background="../ff.jpg">
<table border=0 bgcolor="#FFFFCC" align=center width=98% cellpadding="2" cellspacing="2" bordercolor="#65251C">
  <tr> 
    <td height="30" colspan="7" bgcolor="#FFFFFF"> 
      <p align=center class="p9"><font style="FONT-SIZE: 9pt" color="#FF3333"><marquee><font color="#990099"> 
        <a href="../jh.asp">��^����</a> �w��</font><%=name%><font color="#990099">���{�в����ߡA�ֳ��Q���ۤv���a�A�o�̥i�H�����z���U�حn�D�G(�{�b��X�Фl����o���Ӫ�1/5������)</font></marquee></font></p>
  <tr> 
    <td height="21" colspan="7" bgcolor="#FFFFFF"> 
      <table width="100%" border="0" cellspacing="2" cellpadding="2" bordercolor="#6633FF">
        <tr bgcolor="#FF0000"> 
          <td> 
            <div align="center"><a href=fangwu.asp><font color="#FFFFFF">�@��Ы�</font></a></div>
          </td>
          <td> 
            <div align="center"><a href=fangwu3.asp><font color="#FFFFFF">���Ť��J</font></a></div>
          </td>
          <td> 
            <div align="center"><a href=fangwu4.asp><font color="#FFFFFF">���v��</font></a></div>
          </td>
          <td> 
            <div align="center"><a href=fangwu5.asp><font color="#FFFFFF">���اO��</font></a></div>
          </td>
        </tr>
      </table>
  <tr bordercolor="#000000" bgcolor="#FF0000"> 
    <td height="15" width="14%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">�Ы�����</font></div>
    </td>
    <td height="15" width="9%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">���</font></div>
    </td>
    <td height="15" width="11%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">����</font></div>
    </td>
    <td height="15" colspan="2" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">���P���X</font></div>
    </td>
    <td height="15" width="19%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">�ХD/��Q</font></div>
    </td>
    <td height="15" width="25%" bordercolor="#65251C"> 
      <div align="center"><font size="2" color="#FFFFFF">�ާ@</font></div>
    </td>
  </tr>
  <% 
do while not rs.eof and not rs.bof 
%> 
  <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
    <td width="14%" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%=rs("����")%> 
      </center>
      </font> 
    <td width="9%" class="calen-curr" height="15"><%=rs("���")%> 
      <div align="center"></div>
    <td width="11%" class="ct-def4" height="15"> 
      <div align="center"><font size="2"><%=rs("��m")%></font></div>
    </td>
    <td colspan="2" class="ct-def4" height="15"><%=rs("��D")%> <font color=red><%=rs("id")%></font><font size="2">��</font></td>
    <td width="19%" class="calen-curr" height="15"> 
      <div align="center"><font size="2"><%=rs("��D")%> / </font><font color="ff00ff"><%=rs("��Q")%></font></div>
    </td>
    <td width="25%" class="calen-curr" height="15"><font size="2"> 
      <center>
        <%if rs("��D")<>name and rs("��D")="�L" then%>
        <a href=fangwu1.asp?id=<%=rs("id")%>><font color="0000ff">�ʶR</font></a>
        <%end if%>
        <%if rs("��D")=name then%>
        <a href=fangwu2.asp?id=<%=rs("id")%>><font color=red>��X</font></a>
        <%end if%>
        <%if rs("��D")<>"�L" and rs("��D")<>name then%>
        �өФv�X��
        <%end if%>
      </center>
      </font></td>
  </tr>
  <% 
rs.movenext 
loop 

	set rs=nothing	
	conn.close
	set conn=nothing
%> 
</table>
</html>