<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from �Τ� where id="&info(9),conn
dim page
page=request.querystring("page")
PageSize = 15
myname=request.cookies("Jname")
myname=info(0)
set buys=server.createobject("adodb.recordset")
buys.open "delete * from  ���� where �ɶ�<now()-10",conn,3,3
buys.open "SELECT * FROM ���� order by �ɶ� DESC",conn,3,3
buys.PageSize = PageSize
pgnum=buys.Pagecount
if page="" or clng(page)<1 then page=1
if clng(page) > pgnum then page=pgnum
if pgnum>0 then buys.AbsolutePage=page
%>
<html>

<head>
<title>���ı���</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body bgcolor="#660000" background="../ff.jpg" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<table width="100%" border="1" bgcolor="#990000">
  <tr>
    <td><font color="#FFFFFF">�������ȫ�A�ШӦ��n��(<font color="#FFFF00">3������</font>)�A�_�h��F�A�N�S�����i��F�I 
      �T�ťH�W�ץi�H���ı�����H!�x���H�����i����! �u�����H��´�]����������ȡ��^���H�i�H�Q���ġI �ҥH�U��j�L�٬O�֧֥κɥΤO�a!�W�L�C��Q���ĳ��S���N�H�����A�N�|�ߥX,�P�˶W�L�C��S���H�������H�A�]�|�ߥX�C�C�C�]���O�@�ӤH���න�Ĥ@���A�_�h�h�����ī�A���@�ӤC�Ѥ����Į̄S�N�H�����A�N�|�N�Ҧ��P�˪��W�r���O�����R���^�`�N�G���Ī�����������H�����|���ҷl����</font></td>
  </tr>
</table>
<table width="100%" align="center" cellspacing="2" border="0" cellpadding="5"
bgcolor="#90c088">
  <tr bgcolor="#f7f7f7"> 
    <td align="left"><font size="2">[�@<font color="#990000"><b><%=buys.pagecount%></b></font>��][<font
color="#990000"><%=musers()%></font>�H���ı���] [<a href=shashoulist.asp?page=<%=buys.pagecount-1%>><font color="#0000FF">�W�@��</font></a>][��<%=page%>��][<a href=shashoulist.asp?page= <%=buys.pagecount+1%>><font color="#0000FF">�U�@��</font></a>]</font></td>
  </tr>
  <tr> 
    <td> 
      <table border="0" cellspacing="1" cellpadding="2" width="100%"
bgcolor="#000000" bordercolorlight="#EFEFEF">
        <tr bgcolor="#DFEDFD"> 
          <td width="12%"><font size="2">�ШD���H��</font></td>
          <td width="13%"><font size="2">�Y�N�Q����</font></td>
          <td width="10%"><font size="2">�Q���֪�</font></td>
          <td width="13%"><font size="2">���H�Ī�</font></td>
          <td width="18%"><font size="2">���H����</font></td>
          <td width="21%"><font size="2">�ШD���H�ɶ�</font></td>
          <td width="13%"><font size="2">�O�_���\</font></td>
        </tr>
        <%
count=0
do while not buys.eof and count<buys.PageSize
%>
        <tr bgcolor="#f7f7f7"> 
          <td width="12%"><font size="2"><%=buys("�m�W1")%></font></td>
          <td width="13%"><font size="2"><%=buys("�Q����")%></font></td>
          <td width="10%"><font size="2"><%=buys("�m�W2")%></font></td>
          <td width="13%"><font size="2"><%=buys("���H�Ī�")%></font></td>
          <td width="18%"><font size="2"><%=buys("����")%></font></td>
          <td width="21%"><font size="2"><%=buys("�ɶ�")%> </font> 
          <td width="13%"><font size="2"> 
            
            <a
href="shashouling.asp?name= mmTranslatedValueHiliteColor="HILITECOLOR=%22Dyn%20Untranslated%20Color%22"<%=buys("�Q����")%>"><font color="#FF0000">���</font></a> 
            
            <%if info(2)=10 then %>
            <a
href="shashoulingok.asp?name= mmTranslatedValueHiliteColor="HILITECOLOR=%22Dyn%20Untranslated%20Color%22"<%=buys("�m�W1")%>"><font color="#FF0000">�W�L�ɶ��R��</font></a> 
            <%end if%>
            </font></td>
        </tr>
        <%buys.movenext%>
        <%count=count+1%>
        <%loop
buys.close
rs.close
%>
      </table>
    </td>
  </tr>
</table>

<%
function musers()
dim tmprs
tmprs=conn.execute("Select count(*) As �m�W1 from ����")
musers=tmprs("�m�W1")
set tmprs=nothing
if isnull(musers) then musers=0
end function
rs.open "select ���� from �Τ� where id="&info(9),conn
denji=rs("����")
rs.close
conn.close
set rs=nothing	
set conn=nothing
%><p>&nbsp;</p>
<p align="center"> 
  <%if denji>3 then%>
  <a href="gushashou.htm"><font size="2">�ڭn���ı���</font></a> <a href="javascript:this.location.reload()"><font size="2" >��s</a> 
  <%end if%>
</p>
</body>

</html>