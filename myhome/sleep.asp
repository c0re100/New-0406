<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="homecheck.asp"-->
<html>

<head>
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�鮧</title></head>

<body bgcolor="#000000" background="../ljimage/11.jpg">
<p align="center" style="font-size:16;color:yellow"><font color="#000000"><%=info(0)%>�z�i�H��ߪ��b�a���Τjı�F�G�^~~
</font>
<table border="1" bgcolor="#FFCC99" align="center" width="350" cellpadding="10"
cellspacing="13">
<tr>
<td bgcolor="#BEE0FC">
<table>
<tr>
<td> <font color="#000000"><form method="POST" action="sleep1.asp">
�A�Q�n�𮧡G
<select name="date" size="1">
<option value="0">�s��
<option value="1">�@��
<option value="2">���
<option value="3">�T��
</select>
<select name="time" size="1">
<option value="0">0�p��
<option value="1">1�p��
<option value="2">2�p��
<option value="3">3�p��
<option value="4">4�p��
<option value="5">5�p��
<option value="6">6�p��
<option value="7">7�p��
<option value="8">8�p��
<option value="9">9�p��
<option value="10">10�p��
<option value="11">11�p��
<option value="12">12�p��
<option value="13">13�p��
<option value="14">14�p��
<option value="15">15�p��
<option value="16">16�p��
<option value="17">17�p��
<option value="18">18�p��
<option value="19">19�p��
<option value="20">20�p��
<option value="21">21�p��
<option value="22">22�p��
<option value="23">23�p��
</select>
</font></td>
</tr>
<tr>
<td colspan="2" align="center"><font color="#000000">
<input type="submit" value="�T�w">
<input
type="reset" value="����">
</font></td>
</tr>
<tr>
<td colspan="2" style="font-size:9pt">
<hr>
<font color="#000000"> 1�B�b�a���𮧥i�H�O�@�b���A�åi�H�W�[���O�M�ͩR�ȡF<br>
2�B�����O�I���Ыε��Ť��P�W�[���Ȥ��P�]�Ա��������^�F<br>
3�B�зǽT�p��A�U���ϥαb�����ɶ��I</font></td>
</tr>
<font color="#000000"></form></font>
</table>
</table>

</body>

</html>
