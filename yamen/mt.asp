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
action=request.querystring("action")
if action="�j�a" then action=info(0)
gw=left(action,1)
if IsNumeric(gw)=true then
zz=split(action,"��")
gw=zz(0)
Response.Write "<script language=JavaScript>{alert('�Ǫ��H�����i���I');window.close();}</script>"
	Response.End
end if
if info(0)<>action then
sql="Select  �Ȩ�,grade from �Τ� where �m�W='"& info(0) &"'"
set rs=conn.Execute(sql)
if rs("�Ȩ�")<10000 and rs("grade")<7 then%>
<script language=vbscript>
MsgBox "�A�S���@�U��Ȥl���I�d����F�誺�I"
window.close()
</script>
<%response.end
end if
sql="update �Τ� set �Ȩ�=�Ȩ�-10000 where �m�W='"& info(0) &"'"
Set Rs=conn.Execute(sql)
end if
sql="Select  id,�m�W,�a��,����,�ʧO,�~��,���A,�v��,¾�~,�Ȩ�,���ФH,�|���O,�s��,�D�w,�Z�\,����,����,���O,���s,�y�O,�_��,����,��O,�k�O,����,�t��,�G�B,����,grade,�Y��,ñ�W from �Τ� where �m�W='"& action &"'"
set rs=conn.Execute(sql)
%>
<html>
<head>
<title>�����ɮ�</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
<!--
body, table  { font-size: 9pt; font-family: �s�ө��� }
input        { font-size: 9pt; color: #000000; background-color: #f7f7f7; padding-top: 3px }
.c           { font-family: �s�ө���; font-size: 9pt; font-style: normal; line-height: 12pt;
font-weight: normal; font-variant: normal; text-decoration:
none }
--></style>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#006699" text="#FFFFFF" link="#FFFFCC" vlink="#FFFFCC" alink="#FFFFCC">
<table border="1"
width="429" bgcolor="#FFFFFF" cellspacing="0" cellpadding="1" align="center">
  <tr bgcolor="#333366"> 
    <td colspan="9" height="33"> 
      <div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF">ID:</font><font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("id")%> 
        </font></b></font></font><font color="#FFFFFF"><%=rs("�m�W")%></font></b><font
size="2" color="#FFFFFF">�j�L������I��</font></font></font> </div>
    </td>
  </tr>
  <tr bgcolor="#003366"> 
    <td width="40" height="12"><font size="2" class="c">�m �W:</font></td>
    <td colspan="2" height="12"><%=rs("�m�W")%></td>
    <td height="12" colspan="3">�a��:<%=rs("�a��")%></td>
    <td height="12" width="48"> 
      <div align="center">�W ��:</div>
    </td>
    <td height="12" colspan="2"> <font color="#FFFFCC" size="2"> 
      <%
if info(0)=action then
zddj=-1
else
zddj=rs("����")
end if
if rs("����")<5  then response.write("��ӥE�D")
if rs("����")>=5 and rs("����")<10  then response.write("������")
if rs("����")>=10 and rs("����")<15  then response.write("�p�����N")
if rs("����")>=15 and rs("����")<20  then response.write("�n�W�㻮")
if rs("����")>=20 and rs("����")<35  then response.write("��������")
if rs("����")>=35 and rs("����")<45  then response.write("�@�N�j�L")
if rs("����")>=45 and rs("����")<55  then response.write("����C��")
if rs("����")>=55 and rs("����")<65  then response.write("�D�W�ѤU")
if rs("����")>=65 and rs("����")<75  then response.write("���C�P��")
if rs("����")>=75 and rs("����")<80  then response.write("�w�J�P�D")
if rs("����")>=80 and rs("����")<85  then response.write("�p�P")
if rs("����")>=85 and rs("����")<90  then response.write("�j�P")
if rs("����")>=90 and rs("����")<95  then response.write("���Ҥj�P")
if rs("����")>=95 and rs("����")<100  then response.write("�P�H")
if rs("����")>=100 then response.write("�C�P")
%>
      </font></td>
  </tr>
  <tr bgcolor="#003366"> 
    <td width="40" height="14"><font size="2" class="c">�� �O:</font></td>
    <td height="14" colspan="2"><font color="#FFFFCC" size="2"> 
      <%if rs("�ʧO")="�k"  then response.write("�ӭ�")
if rs("�ʧO")="�k"  then response.write("���k")
%>
      </font></td>
    <td height="14" nowrap width="47">�~��:</td>
    <td height="14" nowrap width="48"><%=rs("�~��")%> ��</td>
    <td height="14" width="40">���A:</td>
    <td height="14" width="48"><%=rs("���A")%></td>
    <td height="14" width="49">�v��:</td>
    <td height="14" width="60"><%=rs("�v��")%></td>
  </tr>
  <tr bgcolor="#003366"> 
    <td width="40">¾�~:</td>
    <td width="32"><%=rs("¾�~")%></td>
    <td colspan="5"><font color="#FFFFFF"></font></td>
    <td width="49">�����G</td>
    <td width="60"><font color="#FFFFFF"><%=rs("����")%><font color="#FFFFCC">��</font></font></td>
  </tr>
</table>
<div align="left"></div>
<table border="1"
width="429" cellspacing="0" cellpadding="1" align="center" bgcolor="#003366">
  <tr> 
    <td colspan="6" height="20"> 
      <div align="center"> <font size="2">�� �� �� ��</font> </div>
    </td>
  </tr>
  <tr> 
    <td width="60" height="25">�{ ���G</td>
    <td height="25" colspan="2"> <%=rs("�Ȩ�")%> �� </td>
    <td width="62" height="25">���ФH�G</td>
    <td height="25" width="52"> <%=rs("���ФH")%> </td>
    <td height="25" width="86">�|�O�G <font color="#FFFFFF"><%=rs("�|���O")%><font color="#FFFFCC">��</font></font> 
    </td>
  </tr>
  <tr> 
    <td width="60" height="20">�s �ڡG</td>
    <td height="24" colspan="2"> <%=rs("�s��")%> �� </td>
    <td width="62" height="24">�D �w�G</td>
    <td height="24" colspan="2"> <%=rs("�D�w")%> </td>
  </tr>
  <tr> 
    <td width="60" height="20">�Z �\�G</td>
    <td height="24" colspan="2"> <%=rs("�Z�\")%> </td>
    <td width="62" height="24">�� ���G</td>
    <td height="24" colspan="2"> <%=rs("����")%> </td>
  </tr>
  <tr> 
    <td width="60" height="20">�� �O�G</td>
    <td colspan="2"> <%=rs("���O")%> </td>
    <td width="62">�� �s�G</td>
    <td colspan="2"> <%=rs("���s")%> </td>
  </tr>
  <tr> 
    <td width="60" height="20">�y �O�G</td>
    <td colspan="2"> <%=rs("�y�O")%> </td>
    <td width="62">�� ���G</td>
    <td width="52"><%=rs("����")%> </td>
    <td width="86">�_��:<%=rs("�_��")%></td>
  </tr>
  <tr> 
    <td width="60" height="20">�� �O�G</td>
    <td width="72"> <%=rs("��O")%> </td>
    <td width="72">�k�O�G<%=rs("�k�O")%></td>
    <td width="62">�� ���G</td>
    <td> <%=rs("����")%> </td>
    
  </tr>
  <tr> 
    <td width="60" height="20">�t ���G</td>
    <td colspan="2"> <%=rs("�t��")%> �G�B�G<%=rs("�G�B")%></td>
    <td width="62">Ĺ / ��G</td>
    <td colspan="2"></td>
  </tr>
  <tr> 
    <td width="60" height="20">������šG</td>
    <td colspan="2"> 
    </td>
    <td width="62">���򵥯šG</td>
    <td width="52"><%=rs("����")%>��</td>
    <td width="86">�޲z����:<%=rs("grade")%>��</td>
  </tr>
  <tr> 
    <td colspan="6" height="199"> 
      <table width="424" border="1" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="196" width="223"> 
            <div align="center"><font color="#000000" size="2"> </font><font color="#000000"><img src="<%=rs("�Y��")%>"></font></div>
          </td>
          <td height="196" width="184" align="left" valign="top">�ӤHñ�W�G<span class="txt"><%=changechr(rs("ñ�W"))%></span></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
</body>

</html>
<%
rs.close
set rs=nothing
conn.close
set conn=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%> 