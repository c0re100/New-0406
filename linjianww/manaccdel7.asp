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
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
page=Request.QueryString("page")
if page="" then page=1
if Not(IsNumeric(page)) then page=1
if page<1 then page=1
page=int(page)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT �m�W,grade,times,allvalue,lasttime,lastip FROM �Τ� WHERE  times=1 AND CDate(lasttime)<now()-29"
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
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5)
show(j)=rs("�m�W")
show(j+1)=rs("grade")
show(j+2)=rs("times")
show(j+3)=rs("allvalue")
show(j+4)=rs("lasttime")
show(j+5)=rs("lastip")
i=i+1
j=j+6
rs.MoveNext
Loop
end if
rs.close
conn.close
set rs=nothing
set conn=nothing%><html>
<head>
<title>�b���޲z</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="readonly/style.css">
</head>
<body bgcolor="#FFFFFF" class=p150 background="0064.jpg">
<div align="center">
  <h1><font color="#FF0000" size="+6">�i�b���޲z�j</font></h1>
<font color="#FF0000">�i�M�z30�Ѱ��ιL�������b���j</font>
<br><a href="manacc.asp">��^</a></div>
<hr noshade size="1" color=009900>
<b> �`�N�ƶ� </b><br>
�H�U <font color=red><%=totalrec%></font> �ӱb�����ιL�����A�B�w�g�W�L30��<%if totalrec>0 then%>�A�O�_�M���H <a href="manaccdel7ok.asp">�O(Yes)</a> <a href="javascript:history.go(-1)">�_(No)</a>
<br>
�p�G�䤤�Y�Ǳb�����Q�M���A�бN�ӱb���K�[�졧<a href="manaccyjadd.asp">�ä[�O�d</a>�����A�Y�i����ӱb���Q�R���C
<hr noshade size="1" color=009900>
<div align=center>
<%for i=1 to totalpage
if page=i then
Response.Write " [" & i & "]"
else
Response.Write " <a href=manaccdel7.asp?page=" & i & ">[" & i & "]</a>"
end if
next%></div>
<hr noshade size="1" color=009900>
<table border="1" cellspacing="0" cellpadding="6" bordercolorlight="#999999" bordercolordark="#FFFFFF" bgcolor="E0E0E0" align="center">
<tr bgcolor="#3399FF">
<td><font color="#FFFFFF">��</font></td>
<td><font color="#FFFFFF">�Τ�W</font></td>
<td><font color="#FFFFFF">����</font></td>
<td><font color="#FFFFFF">����</font></td>
<td><font color="#FFFFFF">�g���</font></td>
<td><font color="#FFFFFF">�̫�ɶ�</font></td>
<td><font color="#FFFFFF">�̫�ע�</font></td>
</tr>
<%for i=1 to UBound(show) step 6%>
<tr>
<td><%=(page-1)*15+(i+5)/6%></td>
<td><%=show(i)%></td>
<td><%=show(i+1)%></td>
<td><%=show(i+2)%></td>
<td><%=show(i+3)%></td>
<td><%=show(i+4)%></td>
<td><%=show(i+5)%></td>
</tr>
<%next%>
</table><%end if%>
</body>
</html> 