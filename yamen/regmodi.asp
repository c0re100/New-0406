<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
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
action=info(0)
sql="Select  �m�W,�a��,����,�ʧO,�~��,���A,�v��,�H�c,¾�~,�Ȩ�,���ФH,�|���O,�s��,�D�w,�Z�\,����,���O,���s,�y�O,����,��O,����,�t��,����,grade,oicq,�ͤ�,�O�@,�Y��,ñ�W from �Τ� where �m�W='"& action &"'"
set rs=conn.Execute(sql)
%>
<html>

<head>
<title>�����ɮ׭ק�</title>
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

<body topmargin="0" leftmargin="0" background="../images/8.jpg">
<form method="POST" action="regmodi1.asp">
  <table border="1" width="429" cellspacing="0" cellpadding="1" align="center">
    <tr> 
      <td colspan="8" height="33" bgcolor="#000000"> 
        <div align="center"> <font size="2" class="c"><font size="3"><b><font color="#FFFFFF"><%=rs("�m�W")%></font></b><font
size="2" color="#FFFFFF">�j�L������I��</font></font></font> </div>
</td>
</tr>
    <tr background="../images/7.jpg"> 
      <td width="36" height="12"><font size="2" class="c">�m �W:</font></td>
      <td colspan="2" height="12"><%=rs("�m�W")%></td>
      <td height="12" colspan="2">�a��: 
        <input type="text"
name="diqu" size="10"
style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("�a��")%>">
</td>
      <td height="12" width="66"> 
        <div align="center">�W ��:</div>
</td>
      <td height="12" colspan="2"> <font color="#0000FF" size="2"> 
        <%if rs("����")<5  then response.write("��ӥE�D")
if rs("����")>=5 and rs("����")<10  then response.write("������")
if rs("����")>=10 and rs("����")<15  then response.write("�p�����N")
if rs("����")>=15 and rs("����")<20  then response.write("�n�W�㻮")
if rs("����")>=20 and rs("����")<35  then response.write("��������")
if rs("����")>=35 and rs("����")<45  then response.write("�@�N�j�L")
if rs("����")>=45 and rs("����")<55  then response.write("����C��")
if rs("����")>=55 and rs("����")<65  then response.write("�D�W�ѤU")
if rs("����")>=65 and rs("����")<75  then response.write("���C�P��")
if rs("����")>=100 then response.write("�C�P")
%>
        </font></td>
</tr>
    <tr background="../images/7.jpg"> 
      <td width="36" height="20"><font size="2" class="c">�� �O:</font></td>
      <td height="20" width="38"><font color="#000000" size="2"> 
        <%if rs("�ʧO")="�k"  then response.write("�ӭ�")
if rs("�ʧO")="�k"  then response.write("���k")
%>
        </font></td>
      <td height="20" width="35">�~��:</td>
      <td height="20" nowrap width="71"> 
        <input type="text"
name="nianling" size="2"
style="font-family: Tahoma; font-size: 12px"
maxlength="3" value="<%=rs("�~��")%>">
�� </td>
      <td height="20" width="40" background="../images/7.jpg">���A:</td>
      <td height="20" width="66"><%=rs("���A")%></td>
      <td height="20" width="36">�v��:</td>
      <td height="20" width="73"><%=rs("�v��")%></td>
</tr>
    <tr background="../images/7.jpg"> 
      <td width="36"><font size="2" class="c">Email:</font></td>
      <td colspan="3" nowrap> 
        <input type="text"
name="email" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="50" value="<%=rs("�H�c")%>">
</td>
      <td width="40">OIcq:</td>
      <td width="66"> 
        <input type="text"
name="oicq" size="8"
style="font-family: Tahoma; font-size: 12px"
maxlength="9" value="<%=rs("oicq")%>">
</td>
      <td width="36">¾�~:</td>
      <td width="73"><%=rs("¾�~")%></td>
</tr>
</table>
<div align="left"></div>
  <table border="1"
width="424" cellspacing="0" cellpadding="1" align="center">
    <tr background="../images/7.jpg"> 
      <td colspan="5" height="20"> 
        <div align="center"> <font size="2">�� �� �� ��</font> </div>
      </td>
    </tr>
    <tr> 
      <td width="52" height="2" background="../images/7.jpg">�{ ���G</td>
      <td width="169" height="2" background="../images/7.jpg"><%=rs("�Ȩ�")%> ��</td>
      <td width="56" height="2" background="../images/7.jpg">���ФH�G</td>
      <td height="2" width="50" background="../images/7.jpg"><%=rs("���ФH")%></td>
      <td height="2" width="84" background="../images/7.jpg">�|���O�G<%=rs("�|���O")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">�s �ڡG</td>
      <td width="169" height="24" background="../images/7.jpg"><%=rs("�s��")%> ��</td>
      <td width="56" height="24" background="../images/7.jpg">�D �w �G</td>
      <td height="24" colspan="2" background="../images/7.jpg"><%=rs("�D�w")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">�Z �\�G</td>
      <td width="169" height="24" background="../images/7.jpg"><%=rs("�Z�\")%></td>
      <td width="56" height="24" background="../images/7.jpg">�� �� �G</td>
      <td height="24" colspan="2" background="../images/7.jpg"><%=rs("����")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">�� �O�G</td>
      <td width="169" background="../images/7.jpg"><%=rs("���O")%></td>
      <td width="56" background="../images/7.jpg">�� �s �G</td>
      <td colspan="2" background="../images/7.jpg"><%=rs("���s")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">�y �O�G</td>
      <td width="169" background="../images/7.jpg"><%=rs("�y�O")%></td>
      <td width="56" background="../images/7.jpg">�� �� �G</td>
      <td colspan="2" background="../images/7.jpg"><%=rs("����")%> </td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">�� �O�G</td>
      <td width="169" background="../images/7.jpg"><%=rs("��O")%></td>
      <td width="56" background="../images/7.jpg">�� �� �G</td>
      <td colspan="2" background="../images/7.jpg"><%=rs("����")%></td>
    </tr>
    <tr> 
      <td width="52" height="20" background="../images/7.jpg">�t ���G</td>
      <td width="169" background="../images/7.jpg"><%=rs("�t��")%></td>
      <td width="56" background="../images/7.jpg">Ĺ / ��G</td>
      <td colspan="2" background="../images/7.jpg"></td>
    </tr>
    <tr> 
      <td width="52" height="2" background="../images/7.jpg">�� ���G</td>
      <td width="169" height="2" background="../images/7.jpg"> �䯫</td>
      <td width="56" height="2" background="../images/7.jpg">�� �šG</td>
      <td width="50" height="2" background="../images/7.jpg"><%=rs("����")%>��</td>
      <td width="84" height="2" background="../images/7.jpg">�޲z��:<%=rs("grade")%>��</td>
    </tr>
    <tr> 
      <td width="52" height="11" background="../images/7.jpg">�� ��G</td>
      <td width="169" height="11" background="../images/7.jpg"> 
        <input type="text"
name="shengri" size="10"
style="font-family: Tahoma; font-size: 12px"
maxlength="10" value="<%=rs("�ͤ�")%>">
      </td>
      <td width="56" height="11" background="../images/7.jpg">�O �@�G</td>
      <td colspan="2" height="11" background="../images/7.jpg"><%=rs("�O�@")%> </td>
    </tr>
    <tr> 
      <td colspan="5" height="199" background="../images/7.jpg"> 
        <table width="421" border="1" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="196" width="237"> 
              <div align="center"><font color="#000000" size="2"> </font> <font color="#000000"><img src="<%=rs("�Y��")%>"></font><br>
                <br>
                �Y��url�a�}�G <br>
                <input type="text"
name="touxiang" size="18"
style="font-family: Tahoma; font-size: 12px"
maxlength="100" value="<%=rs("�Y��")%>">
                <br>
                ��J�Y����url�a�}�A�i�H�ק�ۤv<br>
                ���Y���A�z�i�H�κ����B�h��������<br>
                �Y���ӧ@�ۤv���A�Y�����n�Ӥj�G��ĳ<br>
                ���ΡG200*130�j�p </div>
            </td>
            <td height="196" width="178" align="left" valign="top"> 
              <div align="center">�ӤHñ�W�A�g�W�Q�P�B�ͻ����ܡI<span class="txt"> 
                <textarea name="qianming" cols="30" style="font-family: Tahoma; font-size: 12px" rows="20"><%=rs("ñ�W")%></textarea>
                </span></div>
            </td>
          </tr>
        </table>
        <div align="right"></div>
        <div align="right"></div>
      </td>
    </tr>
  </table>
<div align="center"><br>
<input type="submit" value="�����ƭק�" name="B1" tyle="font-family: Tahoma; font-size: 12px">
</div>
</form>
<div align="center">
<br>
</div>
</body>

</html>
<%
rs.close
set rs=nothing
function changechr(str)
changechr=replace(replace(replace(replace(str,"<","&lt;"),">","&gt;"),chr(13),"<br>")," ","&nbsp;")
changechr=replace(replace(replace(replace(changechr,"[img]","<img src="),"[b]","<b>"),"[red]","<font color=CC0000>"),"[big]","<font size=7>")
changechr=replace(replace(replace(replace(changechr,"[/img]","></img>"),"[/b]","</b>"),"[/red]","</font>"),"[/big]","</font>")
end function
%> 