<%
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<html>
<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<style>input, body, select, td { font-size: 14; line-height: 160% }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�F�C����ȴ�</title></head>
<body LEFTMARGIN="0" TOPMARGIN="0" BGCOLOR="#996699">
<p align="center" style="font-size:16;color:yellow"><font color="#000000">�w��<%=info(0)%>���{���ȴ̥� 
</font> <table border="1" align="center" width="399" cellpadding="10"
cellspacing="13" BACKGROUND="../../IMAGES/BJ31.JPG"> <tr> <td bgcolor="#BEE0FC" height="230" BACKGROUND="../../ljimage/downbg.jpg"> 
<table width="358"> <tr> <td> <form method="POST" action="pub1.asp"> <tr> <td align="center">�A�Q�n�𮧡G 
<select name="time" size="1"> <option value="0">0�p�� <option value="1" selected>1�p�� 
<option value="2">2�p�� <option value="3">3�p�� <option value="4">4�p�� <option value="5">5�p�� 
<option value="6">6�p�� <option value="7">7�p�� <option value="8">8�p�� <option value="9">9�p�� 
<option value="10">10�p�� <option value="11">11�p�� <option value="12">12�p�� <option value="13">13�p�� 
<option value="14">14�p�� <option value="15">15�p�� <option value="16">16�p�� <option value="17">17�p�� 
<option value="18">18�p�� <option value="19">19�p�� <option value="20">20�p�� <option value="21">21�p�� 
<option value="22">22�p�� <option value="23">23�p�� </select></td></tr> <tr> <td colspan="2" align="center"><input type="submit" value="�T�w"> 
<INPUT TYPE="button" VALUE="����" ONCLICK="javascript:window.close();" NAME="button"></td></tr> 
<tr> <td colspan="2" style="font-size:9pt"> <hr> 1�B�b�ȴ̤��𮧥i�H�O�@�b���A�åi�H�W�[���O�M�ͩR�ȡF<br> 
2�B�C1�p�ɦ��A�ȶO10��I�W�[�ͩR10�A���O10�F<br> 3�B�зǽT�p��A�U���ϥαb�����ɶ��I</td></tr> </form> </table></table>
</body>

</html>
