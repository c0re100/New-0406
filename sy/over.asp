<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
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
<link rel="stylesheet" href="../css.css">
</head>
<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF">
<br>
<table width="778" border="0" cellspacing="0" cellpadding="0">
<tr>
<td width="569" valign="bottom"><img src="../images/sy5.gif" width="58" height="23"></td>
<td width="209"><img src="../images/sy1.jpg" width="227" height="83"></td>
</tr>
<tr>
<td background="../images/sy4.gif" width="569" align="right"><img src="../images/sy3.gif" width="94" height="42"></td>
<td background="../images/sy4.gif" width="209"><img src="../images/sy2.jpg" width="227" height="42"></td>
</tr>
</table>
<br>
&nbsp;&nbsp;&nbsp;�ӭު��Ȥ@�� <a href="manage.asp">�ӭ޺޲z</a> 
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
<table width="650" border="1" cellspacing="0" cellpadding="2" align="center" bordercolor="#000000">
<tr>
<td>�� �@��<font color=red><%=Totalcount%></font>�i���ȡA�@<%=TotalPages%>���A�ثe����<%=CurPage%>��</td>
<td align="right"> <a href="addnew.asp"><font color="#FF0000">�K�[����</font></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href=over.asp?owner=<%=owner%>>[��s]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=<%=Curpage-1%>>[�W�@��]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=<%=Curpage+1%>>[�U�@��]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=1>[����]</a>&nbsp;<a href=over.asp?owner=<%=owner%>&Curpage=<%=TotalPages%>>[����]</a>&nbsp;</td>
</tr>
</table>
<br>
<table width=650 border=1 cellspacing=0 cellpadding=0 align="center" bordercolor="#000000" >
<tr height=22>
<td height="25" width="334" align="center"><font color="#FF6600"><b>&nbsp;&nbsp;&nbsp;</b>��
�� �� �D </font></td>
<td align="center" colspan="2" height="25"><font color="#FF6600">�� �` ��</font></td>
<td align="center" height="25" width="100"><font color="#FF6600">�Q �i �� </font></td>
<td align="center" height="25" width="93"><font color="#FF6600">�f �P �� �G</font></td>
</tr>
<%I=0
p=RS.PageSize*(Curpage-1)
do while (Not RS.Eof) and (II<RS.PageSize)
p=p+1%>
<tr>
<td height="25" width="334"><a href=disp.asp?ID=<%=rs("id")%>><%=rs("���D")%></a>
</td>
<td colspan="2" height="25">
<div align="center"><b><%=left(rs("��i"),10)%></b></div>
</td>
<td width="100" height="25">
<div align="center"><b><%=left(rs("�Q�i"),10)%></b></div>
</td>
<td width="93" height="25"><b><%if rs("���G")="N" then%>���f�z<% end if %><%if  rs("���G")="0" then %>������<% end if %><% if  rs("���G")="1" then %>����<% end if %></b></td>
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
<table width="650" border="1" cellspacing="0" cellpadding="3" align="center" bordercolorlight="#000000" bordercolordark="#FFFFFF">
<tr>
<td>
<div align="center">�� �����G <%=CurPage%>/<%=TotalPages%> �A�C���G<%=RS.PageSize%>��
</div>
</td>
<td align="right">
<div align="center">���ơG <a href="over.asp?owner=<%=owner%>&CurPage=<%=StartPageNum-1%>"><<</a>
<% For I=StartPageNum to EndPageNum
if I<>CurPage then %> <a href="over.asp?owner=<%=owner%>&CurPage=<%=I%>">[<%=I%>]</a>
<% else %> <b><%=I%></b> <% end if %> <% Next %> <% if EndPageNum<RS.Pagecount then %>
<a href="over.asp?owner=<%=owner%>&CurPage=<%=EndPageNum+1%>">[��h...]</a>
<%end if%> </div>
</tr>
</table>
<%else%> <br>
<div align="center">�ȮɵL���ȡA��<a href="addnew.asp">�K�[�s����</a>
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
<html></html> 