<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();}</script>"
	Response.End 
end if

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
myname=trim(request.querystring("myname"))
PageSize = 15
rs.open "delete * from ������� where �ɶ�<now()-5 and �覡<>'�O�I�d'",conn,3,3
if myname="" then
rs.open "Select id,���~�W,�֦���,�ƶq,�ﹳ,�ɶ�,����,�Ȩ�,�ﹳ From ������� where �覡='�ذe' Order by �ɶ� DESC",conn,3,3
else
rs.open "Select id,���~�W,�֦���,�ƶq,�ﹳ,�ɶ�,����,�Ȩ�,�ﹳ From ������� where �覡='�ذe' and  (�ﹳ='"& myname &"' or �֦���="& info(9)&") Order by �ɶ� DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>���N§�~��</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>

<body leftmargin="0" topmargin="0" background="../../ff.jpg">
<table border="0" height="24" width="700" cellspacing="0" cellpadding="0"
align="center">
  <tbody>
    <tr bgcolor="#FF0000"> 
      <td height="15" width="79%"><font color="#669966"> <font color="#FFFFFF"><font color="#000000">�`�N10�ѥ�������\���L�����~�ڭ̱N�e��U�����B�z!</font><b><a                    href="zslp.asp?myname=<%=info(0)%>"> 
        </a></b><a href="zslp.asp?myname=<%=info(0)%>">�d��ڪ�§�~</a> 
        </font><font color="#669966"><font color="#CC0000" face="����"><a href="javascript:this.location.reload()">��s������</a></font></font></font></td>
      <td height="15" width="21%"> 
        <div align="right"><font color="#669966"><font color="#FFFFFF"><a href="wupin.asp" target="_blank">�B�z�ۤv���~�I�o��</a></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
  <table width="700" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
<td width="100%">
<table border="0" cellspacing="2" cellpadding="2" width="700"
bordercolorlight="#EFEFEF">
          <tr bgcolor="#FFFFFF">
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> ���~�W</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> �e§�H</font></div>
</td>
<td width="5%" height="16">
<div align="center"><font color="#FF6600"> �ƶq</font></div>
</td>
<td width="8%" height="16">
<div align="center"><font color="#FF6600">�ﹳ</font></div>
</td>
<td width="17%" height="16">
<div align="center"><font color="#FF6600"> �� �� </font></div>
</td>
<td width="29%" height="16">
<div align="center"><font color="#FF6600">�� ��</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> ���</font></div>
</td>
<td height="16" colspan="2">
<div align="center"><font color="#FF6600">�ާ@</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr>
<td width="7%" height="29">
<div align="center"> <font color="#0000FF"><%=rs("���~�W")%></font>
</div>
</td>
<td width="7%" height="29">
<div align="center"> <%=rs("�֦���")%> </div>
</td>
<td width="5%" height="29">
<div align="center"> <%=rs("�ƶq")%> </div>
</td>
<td width="8%" height="29">
<div align="center"><font color="#0000FF"><%=rs("�ﹳ")%></font></div>
</td>
<td width="17%" height="29">
<div align="center"> <%=rs("�ɶ�")%> </div>
</td>
<td width="29%" height="29">
<div align="left"></div>
<%=rs("����")%></td>
<td width="7%" height="29">
<div align="center"> <%=rs("�Ȩ�")%> </div>
</td>
<td width="12%" height="29">
<div align="center"><% if info(2)>=9 then%><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">�R��</font></a><%end if%>
<%if len(rs("����"))>5 then zy=left(rs("����"),4)
if rs("�ﹳ")=info(0) and zy<>"�ۧ@�h��" then%>
<a href="zslp1.asp?id=<%=rs("id")%>"><font color="#0000FF"><b>���¤F</b></font></a>
<%else%>
<font color="#0000FF"><b>no�ڪ�</b></font>
<%end if%></div>
</td>
<td width="8%" height="29">
<%if len(rs("����"))>5 then zy=left(rs("����"),4)

if rs("�ﹳ")=info(0) and zy<>"�ۧ@�h��" then%>
<div align="center"><a href="lpjj.asp?id=<%=rs("id")%>"><font color="#0000FF"><b>�ۧ@�h��</b></font></a></div>
<%else%>
<div align="center"><font color="#0000FF"><b>no�ڪ�</b></font></div>
<%end if%>
</td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
conn.close
set rs=nothing
set conn=nothing

%>
</table>
<table border="0" cellspacing="1" cellpadding="1" width="100%" bordercolorlight="#EFEFEF">
<tr>
<td align="left" width="37%" height="2">[�@<font color="red"><b><%=pa%></b></font>��][<font
color="red"><b><%=mu%></b></font>���Ȱe]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="zslp.asp?page=<%=page-1%>">�W�@��</a>][��<%=page%>��][<a                    href="zslp.asp?page=<%=page+1%>">�U�@��</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As �ƶq from ������� where �覡='�ذe'")
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
end function

%>
</body>
