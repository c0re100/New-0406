<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="data.asp"--><%
sql="SELECT * FROM �ө� where ���D='"&info(0)&"'"
set rs1=conn1.Execute (sql)
if (not rs1.EOF) or (not rs1.BOF) then
set rs1=nothing
conn1.close
set conn1=nothing%>
<script language="vbscript">
alert("�A�w�g�}�]�F�ө��F�I")
window.close()
</script>
<%Response.End 
end if
rs1.close
set rs1=nothing
rs.open "select �|������,����,�Ȩ� FROM �Τ� WHERE id="&info(9),conn
if rs("�|������")=0 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�ө��ѪO�ثe����|���}��I');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
if rs("�Ȩ�")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�A�ݭn�a��2���Ȩ�@���ө����`�U����I');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
if rs("����")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�}�]�ө��ݭn�A���԰��Ŧb100�ťH�W�I');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
rs.close
rs.open "select id FROM ���~ WHERE (����='�A��' or ����='�k��' or ����='����' or ����='�ī~' or ����='�˹�' or ����='����') and  �ƶq>0 and �֦���="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('�藍�_�A�}�]�ө��ݭn�A��6�إH�W���~�I');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<head>
<title>�}�]�ө�</title>
</head>

<body bgcolor="#99CCFF" background="../linjianww/0064.jpg">
<p align="center"><font color="#800080" size="6">�}�]�ө�</font></p>
<form method="POST" action="creatok1.asp">
  <div align="center"> <center> 
      <table border="1" width="70%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr> 
          <td width="39%">�ө��W:</td>
          <td width="66%"> 
            <input type="text" name="shopname" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          </td>
        </tr>
        <tr> 
          <td width="39%">��&nbsp; ��:</td>
          <td width="66%"> 
            <input type="text" name="memo" size="33" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          </td>
        </tr>
        <tr> 
          <td width="39%">�g�綵��:</td>
          <td width="66%"> 
            <select size="1" name="shoptype" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
              <option value="�п��" selected>�п��</option>
              <option value="���~">���~</option>
            </select>
          </td>
        </tr>
        <tr> 
          <td width="105%" colspan="2"> 
            <p align="center"> 
              <input type="submit" value="�� ��" name="B1">
          </td>
        </tr>
      </table>
    </center></div></form>
