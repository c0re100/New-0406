<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();}</script>"
	Response.End 
end if

%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if info(0)="" then Response.Redirect "../error.asp?id=210"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
dim page
page=request.querystring("page")
myname=trim(request.querystring("myname"))
PageSize = 15
rs.open "delete * from ������� where �ɶ�<now()-10 and �覡<>'�O�I�d'",conn,3,3
if myname="" then
	rs.open "Select id,���~�W,�֦���,�ƶq,�ﹳ,�ɶ�,����,�Ȩ�,�G���,�ﹳ From ������� where �覡='���' Order by �ɶ� DESC",conn,3,3
else
	rs.open "Select id,���~�W,�֦���,�ƶq,�ﹳ,�ɶ�,����,�Ȩ�,�G���,�ﹳ From ������� where �覡='���' and (�ﹳ='"& myname &"' or �֦���="& info(9) &") Order by �ɶ� DESC",conn,3,3
end if

rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>�ӶT���</title>
<style type="text/css">td           { font-family: �s�ө���; font-size: 9pt }
body         { font-family: �s�ө���; font-size: 9pt }
select       { font-family: �s�ө���; font-size: 9pt }
a            { color: #FFC106; font-family: �s�ө���; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: �s�ө���; font-size: 9pt; text-decoration:
underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#660000" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table border="0" height="24" width="712" cellspacing="0" cellpadding="0"
align="center">
  <tbody> 
  <tr>
    <td height="15" width="79%" bgcolor="#99CCFF"><font color="#669966"> <font color="#FFFFFF"><b><font color="#0000FF">�ӶT���</font>(</b><font color="#000000">�`�N10�ѥ�������\���L�����~�ڭ̱N�e��U�����B�z!</font><b>�^<a                    href="smjy.asp?myname=<%=info(0)%>"><font color="#990000">�d��ڪ����</font></a> 
      </b></font></font><font face="����"><a href="javascript:this.location.reload()"><font color="#990000">��s������</font></a></font><font color="#990000"><b> 
      </b></font></td>
    <td height="15" width="21%" bgcolor="#99CCFF"> 
      <div align="right"><font color="#669966"><font color="#FFFFFF"><b><a href="wupin.asp" target="_blank"><font color="#990000">�B�z�ۤv���~�I�o��</font></a></b></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
  <table width="712" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
<td width="100%">
        <table border="1" cellspacing="1" cellpadding="0" width="712"
bordercolorlight="#EFEFEF" align="center">
          <tr bgcolor="#FFFFFF">
            <td width="11%" height="16"> 
              <div align="center"><font color="#FF6600"> ���~�W</font></div>
</td>
            <td width="8%" height="16"> 
              <div align="center"><font color="#FF6600"> �֦���</font></div>
</td>
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600"> �ƶq</font></div>
</td>
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600">�ﹳ</font></div>
</td>
            <td width="13%" height="16"> 
              <div align="center"><font color="#FF6600"> �� �� </font></div>
</td>
            <td width="23%" height="16"> 
              <div align="center"><font color="#FF6600">�� ��</font></div>
</td>
            <td width="7%" height="16"> 
              <div align="center"><font color="#FF6600"> ���</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600">�{��</font></div>
</td>
            <td width="13%" height="16"> 
              <div align="center"><font color="#FF6600">�ʶR</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
<tr>
            <td width="11%" height="21"> 
              <div align="center"> <font color="#FFFFCC"><%=rs("���~�W")%></font>
</div>
</td>
            <td width="8%" height="21"> 
              <div align="center"> <%=rs("�֦���")%> </div>
</td>
            <td width="9%" height="21"> 
              <div align="center"> <%=rs("�ƶq")%> </div>
</td>
            <td width="9%" height="21"> 
              <div align="center"><%=rs("�ﹳ")%></div>
</td>
            <td width="13%" height="21"> 
              <div align="center"> <%=rs("�ɶ�")%> </div>
</td>
            <td width="23%" height="21"> 
              <div align="left"></div>
<%=rs("����")%></td>
            <td width="7%" height="21"> 
              <div align="center"> <%=rs("�Ȩ�")%> </div>
</td>
<td width="7%" height="21">
<div align="center"> <%=rs("�G���")%> </div>
</td>
            <td width="13%" height="21"> 
              <div align="center"> 
                <% if info(2)>=9 then%>
                <a href="del.asp?id=<%=rs("id")%>"><font color="#FFFFFF">�R��</font></a> 
                <%end if%>
                <%if rs("�ﹳ")=info(0) then%>
                <a href="smjy1.asp?id=<%=rs("id")%>"><b><font color="#FFFFFF">�n�F</font></b></a> 
                <%else%>
                <font color="#FFFFFF">no���</font> 
                <%end if%>
              </div>
            </td>
</tr>
<%rs.movenext%>
<%count=count+1%>
<%loop
pa=rs.pagecount
mu=musers()
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
        <table border="0" cellspacing="1" cellpadding="1" width="712" bordercolorlight="#EFEFEF" align="center">
          <tr>
<td align="left" width="37%" height="2">[�@<font color="red"><b><%=pa%></b></font>��][<font
color="red"><b><%=mu%></b></font>�����]</td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="smjy.asp?page=<%=page-1%>">�W�@��</a>][��<%=page%>��][<a                    href="smjy.asp?page=<%=page+1%>">�U�@��</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As �ƶq from ������� where �覡='���'")
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>

</body>