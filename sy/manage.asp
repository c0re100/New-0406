<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10  then Response.Redirect "../error.asp?id=439"
dim conn,rs,userconn,users
username=ljjh_nickname
If Request.QueryString("CurPage") = "" or Request.QueryString("CurPage") = 0 then
CurPage = 1
Else
CurPage = CINT(Request.QueryString("CurPage"))
End If
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
%>

<html>
<head>
<title>�ӭު��Ȥ@��</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type=text/css><!--td {  font-family: �s�ө���; font-size: 9pt}body {  font-family: �s�ө���; font-size: 9pt}select {  font-family: �s�ө���; font-size: 9pt}A {text-decoration: none; font-family: "�s�ө���"; font-size: 9pt}A:hover {text-decoration: underline; color: #CC0000; font-family: "�s�ө���"; font-size: 9pt} .big {  font-family: �s�ө���; font-size: 12pt}
--></style>

</head>
<body leftmargin="0" topmargin="0" bgcolor="#3a4b91" background="../linjianww/0064.jpg" text="#000000">
<table width="650" border="0" cellspacing="2" cellpadding="2" align="center" bgcolor="#0066CC">
<tr>
<td colspan="2" height="26">
      <div align="center"><font size="+2" color="#FFFFFF">�ӭު��Ȥ@��</font><font size="+2">&nbsp;</font></div>
</td>
</tr>
</table>

<form action=over.asp method=get>
<%
set rs=server.createobject("adodb.recordset")
rs.open "SELECT * FROM �ӭ�  ORDER BY id DESC",Conn,1,1
if not rs.eof and not rs.bof then
RS.PageSize=15
Dim TotalPages
TotalPages = RS.PageCount

If CurPage>RS.Pagecount Then
CurPage=RS.Pagecount
end if

RS.AbsolutePage=CurPage

rs.CacheSize = RS.PageSize

Dim Totalcount
Totalcount =INT(RS.recordcount)
%>
  <table width="650" border="1" cellspacing="0" cellpadding="2" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#99CCFF">
    <tr>
      <td width="261"><font size="2" color="#000000">�� �@��<%=Totalcount%>�i���ȡA�@<%=TotalPages%>���A�ثe����<%=CurPage%>��</font></td>
      <td align="right" width="375"><font size="2" color="#000000"> &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=manage.asp?owner=<%=owner%>>[��s]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage-1%>>[�W�@��]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=Curpage+1%>>[�U�@��]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=1>[����]</a>&nbsp;<a href=manage.asp?owner=<%=owner%>&Curpage=<%=TotalPages%>>[����]</a>&nbsp;</font></td>
</tr>
</table>
  <table width=650 border=1 cellspacing=0 cellpadding=1 align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF" bgcolor="#FFFFFF">
    <tr height=22 bgcolor="#FF6600"> 
      <td height="25" width="264"><font size="2"><b><font color="#000000">&nbsp;&nbsp;&nbsp;</font></b><font color="#FFFFFF">���ȼ��D</font></font></td>
      <td align="center" colspan="2" height="25"><font size="2" color="#FFFFFF">���`��</font></td>
      <td align="center" height="25" width="91"><font size="2" color="#FFFFFF">�Q�i�� 
        </font></td>
      <td align="center" height="25" width="179"><font size="2" color="#FFFFFF">�f�P���G</font></td>
</tr>
<%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (II<RS.PageSize)
p=p+1%>
    <tr> 
      <td height="25" width="264"><font size="2"><a href=rootdisp.asp?ID=<%=rs("id")%>><%=rs("���D")%></a> 
        </font></td>
<td colspan="2" height="25"><b><font size="2"><%=left(rs("��i"),10)%></font></b></td>
      <td width="91" height="25"><b><font size="2"><%=left(rs("�Q�i"),10)%></font></b></td>
      <td width="179" height="25"><b><font size="2"> 
        <%if rs("���G")="N" then%>
        ���f�z 
        <% end if %>
        <%if  rs("���G")="0" then %>
        ������ 
        <% end if %>
        <% if  rs("���G")="1" then %>
        ���� 
        <% end if %>
        <a href="del.asp?ID=<%=rs("id")%>">�R��</a> </font></b></td>
</tr>
<%rs.movenext
II=II+1
loop%>
</table>
<%StartPageNum=1
do while StartPageNum+15<=CurPage
StartPageNum=StartPageNum+15
Loop

EndPageNum=StartPageNum+14

If EndPageNum>RS.Pagecount then EndPageNum=RS.Pagecount
%>
  <table width="650" border="1" cellspacing="0" cellpadding="3" align="center" bgcolor="#99CCFF" bordercolorlight="#000000" bordercolordark="#FFFFFF">
    <tr>
<td>
<div align="center">�� �����G <font color="#CC0000"><%=CurPage%></font>/<%=TotalPages%>
�A�C���G<font color="#CC0000"><%=RS.PageSize%></font>�� </div>
</td>
<td align="right">
<div align="center">���ơG <a href="manage.asp?owner=<%=owner%>&CurPage=<%=StartPageNum-1%>"><<</a>
<% For I=StartPageNum to EndPageNum
if I<>CurPage then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=I%>">[<%=I%>]</a>
<% else %> <font color="#CC0000"><b><%=I%></b></font> <% end if %> <% Next %>
<% if EndPageNum<RS.Pagecount then %> <a href="manage.asp?owner=<%=owner%>&CurPage=<%=EndPageNum+1%>">[��h...]</a>
<%end if%> </div>
</tr>
</table>
<%else%> <br>
  <div align="center"><font color="#000000">�ȮɵL���ȡA��</font><font color="#FF0001"><a href="index.asp">�K�[�s����</a></font> 
    <%
end if%>
  </div>
</form>


<div align="center"> </div>
</body>
</html>
<%
rs.close
set rs=nothing
%> 