<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="data.asp"--><%
sql="SELECT �ө��W,���D,�`�U���,id FROM �ө� where ���D='"&info(0)&"'"
rs1.open sql,conn1,1,1
if rs1.EOF or rs1.BOF then Response.Redirect "../error.asp?id=484"
dim shopname
shopname=rs1("�ө��W")
name=rs1("���D")
shopmoney=rs1("�`�U���")
id=rs1("id")
if name<>info(0) then
set rs1=nothing
conn1.close
set conn1=nothing
%>
<script language="vbscript">
alert("�o�a�ө����O�A�}���I")
history.back(1)
</script>
<%Response.End 
end if
%>
<head>
<title><%=shopname%></title>
</head>

<body bgcolor="#0099CC" BACKGROUND="../linjianww/0064.jpg" text="#000000" topmargin="0">
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="color: #E0E4E0; font-size: 10pt">
  <tr bgcolor="#0033CC"> 
    <td colspan="9" align="center"> <font size="+3" color="#FFFFFF"> 
      <%rs1.close
	  sql="select * from ���~ where �֦���="&id&" and ����<>'�d��'  and �˳�=false and �ƶq>0"
rs.Open sql,conn,1,1
if not rs.EOF then
%>
      ���~��</font> </td>
  </tr>
  <tr bgcolor="#0099FF"> 
    <td width="21%" align="center"><font color="#FFFFFF">�ӫ~�W</font></td>
    <td width="12%" align="center"><font color="#FFFFFF">����</font></td>
    <td width="10%" align="center"><font color="#FFFFFF">����</font></td>
    <td width="11%" align="center"><font color="#FFFFFF">�ƶq</font></td>
    <td width="46%" align="center"><font color="#FFFFFF">�] �m</font></td>
  </tr>
  <%while not rs.EOF or rs.BOF %>
  <tr> 
    <td width="21%" align="center" background="bg.gif"><font color="#000000"><%=rs("���~�W")%></font></td>
    <td width="12%" align="center" background="bg.gif"><font color="#000000"><%=rs("����")%></font></td>
    <td width="10%" align="center" background="bg.gif"><font color="#000000"><%=rs("�Ȩ�")%></font></td>
    <td width="11%" align="center" background="bg.gif"><font color="#000000"><%=rs("�ƶq")%></font></td>
    <td width="46%" align="center" background="bg.gif"> 
      <form method="POST" action="modify.asp">
        <font color="#000000"> 
        <input type="hidden" name="name" value=<%=rs("���~�W")%>>
        <input type="hidden" name=aaa value=<%=rs("�֦���")%>>
        <input type="submit" value="�[�J�ө�" name="submit">
        </font> 
      </form>
    </td>
  </tr>
  <% 
rs.MoveNext 
wend 
end if
%>
</table>
<table border="1" width="100%" cellspacing="0" cellpadding="0" style="color: #E0E4E0; font-size: 10pt">
  <tr bgcolor="#0033CC"> 
    <td colspan="9" align="center"> <font color="#FFFFFF" size="+3">�ө����~��</font></td>
  </tr>
  <tr bgcolor="#0099FF"> 
    <td width="21%" align="center"><font color="#FFFFFF">�ӫ~�W</font></td>
    <td width="12%" align="center"><font color="#FFFFFF">����</font></td>
    <td width="10%" align="center"><font color="#FFFFFF">����</font></td>
    <td width="11%" align="center"><font color="#FFFFFF">�ƶq</font></td>
    <td width="46%" align="center"><font color="#FFFFFF">�ջ�</font></td>
  </tr>
  <%rs.close
  set rs=nothing
conn.close
set conn=nothing
sql="select * from �ө����~ where �֦���="&id&" and �˳�=false and �ƶq>0"
rs1.Open sql,conn1,1,1
if not rs1.EOF then

  while not rs1.EOF or rs1.BOF %>
  <tr> 
    <td width="21%" align="center" background="bg.gif"><font color="#000000"><%=rs1("���~�W")%></font></td>
    <td width="12%" align="center" background="bg.gif"><font color="#000000"><%=rs1("����")%></font></td>
    <td width="10%" align="center" background="bg.gif"><font color="#000000"><%=rs1("�Ȩ�")%></font></td>
    <td width="11%" align="center" background="bg.gif"><font color="#000000"><%=rs1("�ƶq")%></font></td>
    <td width="46%" align="center" background="bg.gif"> 
      <form method="POST" action="modifyshang.asp">
        <font color="#000000"> 
        <input type="hidden" name="name2" value=<%=rs1("���~�W")%>>
        <input type="hidden" name=aaa2 value=<%=rs1("�֦���")%>>
        <select name="select" style="border: 1px solid; background-color:#EEFFEE;font-size: 9pt; border-color:
#000000 solid" >
          <option value="a" selected>���H2</option>
          <option value="b">���H4</option>
          <option value="c">���H6</option>
          <option value="d">���H8</option>
          <option value="e">���H10</option>
          <option value="f">���H50</option>
          <option value="g">���H100</option>
          <option>---------------</option>
          <option value="h">���H2</option>
          <option value="i">���H4</option>
          <option value="j">���H6</option>
          <option value="k">���H8</option>
          <option value="l">���H10</option>
          <option value="m">���H50</option>
          <option value="n">���H100</option>
        </select>
        <input type="submit" value="O K" name="submit2">
        </font> 
      </form>
    </td>
  </tr>
  <% 
rs1.MoveNext 
wend 
set rs1=nothing
conn1.close
set conn1=nothing

end if
%>
</table>
<p align="center"><A HREF="modifyshop.asp">��^</A> <BR>
   