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
	rs.open "Select id,���~�W,�֦���,�ƶq,�ɶ�,����,�Ȩ�,�G��� From ������� where �覡='�G��'  Order by �ɶ� DESC",conn,3,3
else
	rs.open "Select id,���~�W,�֦���,�ƶq,�ɶ�,����,�Ȩ�,�G��� From ������� where �覡='�G��'  and  �֦���="& info(9) &" Order by �ɶ� DESC",conn,3,3
end if
rs.PageSize = PageSize
pgnum=rs.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then rs.AbsolutePage=page
%>
<html>
<head>
<title>�G����</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#660000" background="../../ff.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<table border="0" height="24" width="713" cellspacing="0" cellpadding="0"
align="center" bgcolor="#FF0000">
  <tbody> 
  <tr>
      <td height="15" width="80%"><font color="#669966"> <font color="#FFFFFF"><b>(</b><font color="#FFFFCC">�`�N10�ѥ�������\���L�����~�ڭ̱N�e��U�����B�z!</font><b>�^</b></font><font color="#669966"><font color="#FFFFFF"><a  href="esjy.asp?myname=<%=info(0)%>">�d��ڪ��G�⪫�~</a></font><font color="#669966"><font color="#CC0000" face="����"><a href="javascript:this.location.reload()"> 
        </a></font></font></font></font></td>
      <td width="20%" height="15" bgcolor="#FF0000"> 
        <div align="right"><font color="#669966"><font color="#FFFFFF"><a href="wupin.asp" target="_blank">�B�z�ۤv���~�I�o��</a></font></font></div>
</td>
</tr>
</tbody>
</table>
<div align="center">
  <table width="713" align="center" cellspacing="0" border="0"
cellpadding="0">
    <tr>
      <td width="100%"> 
        <table border="0" cellspacing="2" cellpadding="2" width="713"
bordercolorlight="#EFEFEF" align="center">
          <tr bgcolor="#FF0000"> 
            <td width="10%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                �� �~ �W</font></div></td>
            <td width="7%" height="11"> 
              <div align="center"><font color="#FFFFFF">�֦��� 
                </font></div></td>
            <td width="7%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                �ƶq</font></div></td>
            <td width="16%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                �� �� </font></div></td>
            <td width="25%" height="11"> 
              <div align="center"><font color="#FFFFFF">�� 
                ��</font></div></td>
            <td width="8%" height="11"> 
              <div align="center"><font color="#FFFFFF"> 
                ��l����</font></div></td>
            <td width="8%" height="11"> 
              <div align="center"><font color="#FFFFFF">�G�����</font></div></td>
            <td width="19%" height="11"> 
              <div align="center"><font color="#FFFFFF">�ʶR</font></div></td>
          </tr>
          <%
count=0
do while not (rs.eof or rs.bof) and count<rs.PageSize
%>
          <tr bgcolor="#006699"> 
            <td width="10%" height="7"> 
              <div align="center"> <font color="#FFFFCC"><%=rs("���~�W")%></font> </div></td>
            <td width="7%" height="7"> 
              <div align="center"> <%=rs("�֦���")%> </div></td>
            <td width="7%" height="7"> 
              <div align="center"> <%=rs("�ƶq")%> </div></td>
            <td width="16%" height="7"> 
              <div align="center"> <%=rs("�ɶ�")%> </div></td>
            <td width="25%" height="7"> 
              <div align="left"></div>
              <%=rs("����")%></td>
            <td width="8%" height="7"> 
              <div align="center"> <%=rs("�Ȩ�")%> </div></td>
            <td width="8%" height="7"> 
              <div align="center"> <%=rs("�G���")%> </div></td>
            <td width="19%" height="7"> 
              <div align="center"><b> 
                <% if info(2)>=9 then%>
                </b><a href="del.asp?id=<%=rs("id")%>"><font color="#FFFF00">�R��</font></a><b> 
                <%end if%>
                </b><a href="esjy1.asp?id=<%=rs("id")%>"><font color="#FFFF00">�n�F</font></a><b> 
                </b></div></td>
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
        <table border="0" cellspacing="1" cellpadding="1" width="713" bordercolorlight="#EFEFEF" align="center">
          <tr>
            <td align="left" width="37%" height="2"><font color="#FFCC00">[�@</font><font color="red"><b><%=pa%></b></font><font color="#FFCC00">��][<b><%=mu%></b>�����]</font></td>
<td align="right" width="63%" height="2">
<div align="center">[<a href="esjy.asp?page=<%=page-1%>">�W�@��</a>][��<%=page%>��][<a href="esjy.asp?page=<%=page+1%>">�U�@��</a>]</div>
</td>
</tr>
</table>
</table>
</div>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As �ƶq from ������� where �覡='�G��'")
musers=tmprs("�ƶq")
set tmprs=nothing
if isnull(musers) then musers=0
end function
%>
</body>