<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"
page=Request.QueryString("page")
if page="" then page=1
if Not(IsNumeric(page)) then page=1
if page<1 then page=1
page=int(page)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT �m�W,grade,times,allvalue,regip,lastip,lasttime FROM �Τ�  ORDER BY allvalue DESC"
rs.open sql,conn,1,1
rs.PageSize=15
totalrec=rs.RecordCount
totalpage=rs.PageCount
if page>totalpage then page=totalpage
if totalrec>0 then
rs.AbsolutePage=page
p=1+(page-1)*rs.PageSize
Dim show()
i=0
j=1
Do while (Not rs.Eof) and (i<rs.PageSize)
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5),show(j+6)
show(j)=rs("�m�W")
show(j+1)=rs("grade")
show(j+2)=rs("times")
show(j+3)=rs("allvalue")
show(j+4)=rs("regip")
show(j+5)=rs("lastip")
show(j+6)=rs("lasttime")
i=i+1
j=j+7
rs.MoveNext
Loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing%>
<html>
<head>
<title>�b���޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head><style type=text/css>
<!--
body,table {font-size: 9pt; font-family: �s�ө���}
.c {  font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt; font-weight: 

normal; font-variant: normal; text-decoration: none}
--></style>
<body bgcolor="#FFFFFF" class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">�i�b���޲z�j</font></h1>
<font color="#FF0000">�i���ֿn���C�X�b���j</font> <br>
<a href="manacc.asp">��^</a></div>
<hr noshade size="1" color=009900>
<b> �`�N�ƶ� </b><br>
�@��� <font color=red><%=totalrec%></font> �ӱb���C<%if totalrec>0 then%>
<hr noshade size="1" color=009900>
<div align=center>
<%for i=1 to totalpage
if page=i then
Response.Write " [" & i & "]"
else
Response.Write " <a href=manaccqueryvalue.asp?page=" & i & ">[" & i & "]</a>"
end if
next%></div>
<hr noshade size="1" color=009900>
<table border="1" cellspacing="0" cellpadding="6" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="E0E0E0" align="center" width="100%">
<tr bgcolor="#3399FF">
<td width="6%" height="23">
<div align="center"><font color="#FFFFFF" size="2">��</font></div>
</td>
<td width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">�Τ�W</font></div>
</td>
<td width="10%" height="23">
<div align="center"><font color="#FFFFFF" size="2">����</font></div>
</td>
<td width="10%" height="23">
<div align="center"><font color="#FFFFFF" size="2">����</font></div>
</td>
<td bgcolor="FF6666" width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">�ֿn��</font></div>
</td>
<td width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">�`�U�ע�</font></div>
</td>
<td width="13%" height="23">
<div align="center"><font color="#FFFFFF" size="2">�̫�ע�</font></div>
</td>
<td width="22%" height="23">
<div align="center"><font color="#FFFFFF" size="2">�̫�ɶ�</font></div>
</td>
</tr>
<%for i=1 to UBound(show) step 7%>
<tr>
<td width="6%"><%=(page-1)*15+(i+6)/7%></td>
<td width="13%"><%=show(i)%></td>
<td width="10%"><%=show(i+1)%></td>
<td width="10%"><%=show(i+2)%></td>
<td width="13%"><%=show(i+3)%></td>
<td width="13%"><%=show(i+4)%></td>
<td width="13%"><%=show(i+5)%></td>
<td width="22%"><%=show(i+6)%></td>
</tr>
<%next%>
</table>
<%end if%>
</body>
</html> 