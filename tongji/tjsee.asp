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
<title>通緝管理</title>
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
  <font color=yellow size="6" face="隸書">通緝令</font> </p>
<p>罵人多次者通輯、刷屏多次者通輯、連續殺人多次者通輯，被殺者連續被殺多次殺人者通輯</p>
<p> 
<font size="2" color="#000000"> 
<% 
dim sql 
dim rs 
dim filename 
sql="select top 40 * from 通緝令 where 批準人員<>'未知' order by 請求時間 desc" 
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
 if rs.eof and rs.bof then 


       response.write "<p align='center'><font color=#FFFFFF> 還 沒 有 任 何 記 錄</font> </p>" 
   else 
%>  
</font>  
<div align="center">  
  <table  cellpadding="1" cellspacing="0" border="1" align="center" width=640
bordercolorlight="#000000" bordercolordark="#FFFFFF" style=font-size:9pt>
    <tr> 
      <td align=center>通緝人犯</td>
      <td align=center>請求人員</td>
      <td align=center>通緝原因</td>
      <td align=center>請求附言</td>
      <td align=center>請求時間</td>
      <td align=center>審批</td>
      <td align=center>批準時間</td>
      <td align=center>批準人員</td>
    </tr>
    <%do while not rs.eof%>
    <tr> 
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("通緝人犯")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("請求人員")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("通緝原因")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("請求附言")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("請求時間")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("站長審批")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("批準時間")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"><%=rs("批準人員")%> </font> 
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
