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
PageSize = 15
rs.open "delete * from ������� where �ɶ�<now()-5 and �覡<>'�O�I�d'",conn,3,3
rs.open "Select id,���~�W,�֦���,�ƶq,�ɶ�,���� From ������� where �覡='�O�I�d' and �֦���="& info(9) &" Order by �ɶ� DESC",conn,3,3
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>

<head>
<title>�ڪ��O�I�d</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#660000" background="../../chat/bk.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
  <table width="710" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
<td width="100%">
        <table border="0" cellspacing="1" cellpadding="0" width="710"
bordercolorlight="#EFEFEF" height="8" align="center">
          <tr bgcolor="#FFFFFF">
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600"> ���~�W</font></div>
</td>
<td width="7%" height="16">
<div align="center"><font color="#FF6600"> �֦���</font></div>
</td>
            <td width="6%" height="16"> 
              <div align="center"><font color="#FF6600"> �ƶq</font></div>
</td>
            <td width="21%" height="16"> 
              <div align="center"><font color="#FF6600"> �� �� </font></div>
</td>
            <td width="36%" height="16"> 
              <div align="center"><font color="#FF6600">�� ��</font></div>
</td>
            <td width="9%" height="16"> 
              <div align="center"><font color="#FF6600"> ���X�ƶq</font></div>
</td>
            <td width="12%" height="16"> 
              <div align="center"><font color="#FF6600">�ާ@</font></div>
</td>
</tr>
<%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
          <tr bgcolor="#990000"> 
            <td width="9%" height="26"> 
              <div align="center"> <font color="#FFFFCC"><%=rs("���~�W")%></font>
</div>
</td>
            <td width="7%" height="26"> 
              <div align="center"> <font color="#FFFFFF"><%=rs("�֦���")%></font></div>
</td>
            <td width="6%" height="26"> 
              <div align="center"> <font color="#FFFFFF"><%=rs("�ƶq")%></font></div>
</td>
            <td width="21%" height="26"> 
              <div align="center"><font color="#FFFFFF"><%=rs("�ɶ�")%></font></div>
</td>
            <td width="36%" height="26"> 
              <div align="left"></div>
<font color="#FFFFFF"><%=rs("����")%></font></td>
<form method=POST action='mybxg1.asp?id=<%=rs("id")%>'>
              <td width="9%" height="26"> 
                <div align="center"> <font color="#FFFFFF">
                <select name="wpsl">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select>
                </font></div>
</td>
              <td width="12%" height="26"> 
                <div align="center"><input type="SUBMIT" name="Submit"  value="�i��"></div>
</td>
</form>
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
        <table border="0" cellspacing="1" cellpadding="1" width="710" bordercolorlight="#EFEFEF" align="center">
          <tr bgcolor="#FFCC00"> 
            <td width="37%" height="2" align="left"><font color="#FFFFFF">[�@<b><%=pa%></b>��][<b><%=mu%></b>�Ӫ��~]</font></td>
            <td width="63%" height="2" align="right"> 
              <div align="center"><font color="#FFFFFF">[<a href="mybxg.asp?page=<%=page-1%>">�W�@��</a>][��<%=page%>��][<a href="mybxg.asp?page=<%=page+1%>">�U�@��</a>] 
                <a href="wupin.asp" target="_blank">�i�����~�s�J�O�I�d�j</a> </font></div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As �ƶq from ������� where �覡='�O�I�d' and �֦���="& info(9))
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>
</body>
