<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")%>
<!--#include file="data1.asp"-->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�q�r�޲z</title>
<link rel="stylesheet" href="../setup.css">
</head>
<%
   const MaxPerPage=500
   dim totalPut   
   dim CurrentPage
   dim TotalPages
   dim i,j
%>

<body text="#FFFFFF" topmargin="0" bgcolor="#660000"><div align="center"><center> 
<p><br>
  <font color=yellow size="6" face="����">�q�r�O</font> </p>
<p>�|�H�h���̳q��B��̦h���̳q��B�s����H�h���̳q��A�Q���̳s��Q���h�����H�̳q��</p>
<p> 
<font size="2" color="#000000"> 
<% 
dim sql 
dim rs 
dim filename 
sql="select top 40 * from �q�r�O where ��ǤH��<>'����' order by �ШD�ɶ� desc" 
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
 if rs.eof and rs.bof then 


       response.write "<p align='center'><font color=#FFFFFF> �� �S �� �� �� �O ��</font> </p>" 
   else 
%>  
</font>  
<div align="center">  
  <table  cellpadding="1" cellspacing="0" border="1" align="center" width=640
bordercolorlight="#000000" bordercolordark="#FFFFFF" style=font-size:9pt>
    <tr> 
      <td align=center>�q�r�H��</td>
      <td align=center>�ШD�H��</td>
      <td align=center>�q�r��]</td>
      <td align=center>�ШD����</td>
      <td align=center>�ШD�ɶ�</td>
      <td align=center>�f��</td>
      <td align=center>��Ǯɶ�</td>
      <td align=center>��ǤH��</td>
    </tr>
    <%do while not rs.eof%>
    <tr> 
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("�q�r�H��")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("�ШD�H��")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("�q�r��]")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("�ШD����")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("�ШD�ɶ�")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("�����f��")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("��Ǯɶ�")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("��ǤH��")%> </font> 
      </td>
    </tr>
    <%  
	       rs.movenext 
	filename=filename+1 
        if filename>499 then Exit Do 
	   loop 
end if 
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
		  %>
  </table> 
</div> 
<p>&nbsp;</p>
