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

<body text="#FFFFFF" topmargin="0" bgcolor="#0099FF"><div align="center"><center> 
<br>
<font color=yellow size="6" face="隸書">通緝記錄</font> 

<p> 
<font size="2" color="#000000"> 
<% 
dim sql 
dim rs 
dim filename 
sql="select top 40 * from 通緝令 order by 請求時間 desc" 
Set rs= Server.CreateObject("ADODB.Recordset") 
rs.open sql,conn,1,1 
 if rs.eof and rs.bof then 


       response.write "<p align='center'> 還 沒 有 任 何 記 錄 </p>" 
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
      <td align=center>站長操作</td>
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
      <td align="center"><font size="2" color="#FFCC99"><%=rs("批準人員")%> </font> 
      </td>
      <td align="center"><font size="2" color="#FFFFCC"> 
        <%if info(2)>=10 and rs("站長審批")="待審" then%>
        <a href='pizhun.asp?tjname=<%=rs("通緝人犯")%>'>批準</a> 
        <%end if%>
        | 
        <%if info(2)>=10 then%>
        <a href='deltj.asp?tjname=<%=rs("通緝人犯")%>'>刪除</a> 
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
    <p><font size="2" color="#FFFFFF">聊管可以在這裡發布通緝請求，但必須經過站長確認通緝請求在生效</font></p>
    <center>
      <p><font size="2" color="#FFFFFF">通緝人犯 
        <input type="text" name="tjname" size="10" maxlength="10"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
        通緝原因
        <input type="text" name="yuanyin" size="10" maxlength="1000"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
        請求人附言
        <input type="text" name="fuyan" size="10" maxlength="1000"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
        <input type="submit" value="添加" name="B1" class="p9">
        <input type="reset" name="Submit" value="清空">
        </font><font size="2" color="#000000"><br>
      </font>
      </p>
    </center>
  </div>
</form>
