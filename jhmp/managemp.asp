<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=Request.QueryString("ID")
if info(6)<>"掌門" then Response.Redirect "../error.asp?id=451"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or inStr(id,"&")<>0 or inStr(id,"&&")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../error.asp?id=54"
rs.open "select 同盟 FROM 用戶 WHERE id=" & info(9),conn
tongmen=rs("同盟")
rs.close
rs.open "Select 門派,簡介,掌門,適合 from 門派 where 門派='"&Id&"'",conn
if rs.eof or rs.bof then
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
			Response.Write "<script language=javascript>alert('無此門派！！');location.href = 'javascript:history.back()';</script>"
			response.end
end  if
%>
<html>
<head>
<title>門派內容修改</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css>
<!--
body,table {font-size: 9pt; font-family: 新細明體}
input {  font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px}
.c {  font-family: 新細明體; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: normal; font-variant: normal; text-decoration: none}
--></style>
</head>
<body bgcolor="#660000" text="#FFFFFF" link="#FFFF00" alink="#FFFF00" vlink="#FFFF00">
<div align="center">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="498" height="31">
<tr align="center" bgcolor="#336633">
      <td width="100%" height="25" bgcolor="#99CCFF"><b><font color="#FF0000" size="4">門派管理</font></b></td>
</tr>
</table>


</div>
<form action="updatmp.asp?subject=<%=rs("門派")%>" method=POST>
  <ul>
    <table border=1 cellspacing=0 cellpadding=1 align="center" width="500" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
      <tr> 
        <td width="180"><font size="2" class="c" color="#FF6600">門派</font></td>
        <td width="308"> <font color="#FF6600"><%=info(5)%> </font></td>
      </tr>
      <tr> 
        <td width="180"><font color="#FF6600">簡介</font></td>
        <td width="308"> <font color="#FF6600"> 
          <input name="comment" value="<%=RS("簡介")%>" size=40 maxlength=50>
          </font></td>
      </tr>
      <tr> 
        <td width="180"><font class="c" size="2" color="#FF6600">掌門&nbsp;</font></td>
        <td width="308"> <font color="#FF6600"><%=rs("掌門")%></font></td>
      </tr>
      <tr> 
        <td width="180"><font color="#FF6600">適合(男/女/所有)</font></td>
        <td width="308"> <font color="#FF6600"> 
          <input name="shihe" value="<%=RS("適合")%>" size=40 maxlength=100>
          </font></td>
      </tr>
      <tr> 
        <td width="180"><font color="#FF6600">掌門禪位（本派弟子）</font></td>
        <td width="308"><font color="#FF6600"> 
          <input name="zhangmen" size=10 maxlength=10 value="無">
          </font></td>
      </tr>
      <tr>
        <td width="180"><font color="#FF6600">同盟幫派</font></td>
        <td width="308"><font color="#FF6600" size="2"><%=tongmen%></font></td>
      </tr>
    </table>
    <div align="center"> 
      <p><font size="3" class="c" color="#000000"> <br>
        </font></p>
      <p><font size="3" class="c">一個人隻能當一個門派的掌門，不能為多家的啊，以最新的設置為準。</font><font size="3" class="c" color="#000000"><br>
        <input type="HIDDEN" name="action" value="RegSubmit">
        <input type="SUBMIT" name="Submit" value="更新" style="border: 2px solid;background-color:#FCCFDF; font-size: 9pt; border-color:
#000000 solid">
        </font></p>
    </div>
  </ul>
</form>
<%rs.close
rs.open "SELECT 姓名 FROM 用戶 where 門派='"&info(5)&"' and 身份='副掌門'",conn
%>

<table width="100%" border="0">
  <tr>副掌門：
  <%
do while not rs.bof and not rs.eof
%>  
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("姓名")%>' title="取消身份"><%=rs("姓名")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<%rs.close
rs.open "SELECT 姓名 FROM 用戶 where 門派='"&info(5)&"' and 身份='長老'",conn
%>
<table width="100%" border="0">
  <tr>長老： 
    <%
do while not rs.bof and not rs.eof
%>
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("姓名")%>' title="取消身份"><%=rs("姓名")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<%rs.close
rs.open "SELECT 姓名 FROM 用戶 where 門派='"&info(5)&"' and 身份='護法'",conn
%>
<table width="100%" border="0">
  <tr>護法： 
    <%
do while not rs.bof and not rs.eof
%>
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("姓名")%>' title="取消身份"><%=rs("姓名")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<%rs.close
rs.open "SELECT 姓名 FROM 用戶 where 門派='"&info(5)&"' and 身份='堂主'",conn
%>
<table width="100%" border="0">
  <tr>堂主： 
    <%
do while not rs.bof and not rs.eof
%>
    <td>|<font color="#FFFFFF"><a href='mpdd.asp?you=<%=rs("姓名")%>' title="取消身份"><%=rs("姓名")%></a></font>|</td>
    <%
rs.movenext
loop
%>
  </tr>
</table>
<div align="center">
  <p>&nbsp;</p>
  <p>&nbsp;<input type=button value=關閉窗口 onClick='window.close()' name="button" style="border: 2px solid;background-color:#FCCFDF; font-size: 9pt; border-color:
#000000 solid"></p>
</div>
</body>
</html>
<%
rs.close
			set rs=nothing
			conn.close
			set conn=nothing
%> 