<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<6 then Response.Redirect "../error.asp?id=461" end if%>
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

<body text="#FFFFFF" topmargin="0" bgcolor="#0099FF"><div align="center"><center> 
<br>
<font color=yellow size="6" face="����">�q�r�O��</font> 

<p> 
<font size="2" color="#000000"> 
<% 
dim sql 
dim rs 
dim filename 
sql="select top 40 * from �q�r�O order by �ШD�ɶ� desc" 
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
 if rs.eof and rs.bof then 


       response.write "<p align='center'> �� �S �� �� �� �O �� </p>" 
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
      <td align=center>�����ާ@</td>
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
      <td align="center"><font size="2" color="#FFCC99"><%=rs("��ǤH��")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"> 
        <%if info(2)>=10 and rs("�����f��")="�ݼf" then%>
        <a href='pizhun.asp?tjname=<%=rs("�q�r�H��")%>'>���</a> 
        <%end if%>
        | 
        <%if info(2)>=10 then%>
        <a href='deltj.asp?tjname=<%=rs("�q�r�H��")%>'>�R��</a> 
        <%end if%>
        </font> </td>
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
<form method="POST" action="tjok.asp">
  <div align="center"> 
    <p><font size="2" color="#FFFFFF">��ޥi�H�b�o�̵o���q�r�ШD�A�������g�L�����T�{�q�r�ШD�b�ͮ�</font></p>
    <center>
      <p><font size="2" color="#FFFFFF">�q�r�H�� 
        <input type="text" name="tjname" size="10" maxlength="10"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
        �q�r��]
        <input type="text" name="yuanyin" size="10" maxlength="1000"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
        �ШD�H����
        <input type="text" name="fuyan" size="10" maxlength="1000"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
        <input type="submit" value="�K�[" name="B1" class="p9">
        <input type="reset" name="Submit" value="�M��">
        </font><font size="2" color="#000000"><br>
      </font>
      </p>
    </center>
  </div>
</form>
