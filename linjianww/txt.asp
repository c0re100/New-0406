<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<html>
<head>
<title>�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>
<body bgcolor="#FFFFFF" background="0064.jpg">
<p><font size="5" face="����_GB2312" color="#3333FF">�w��z�ϥ��F�C����޲z�t�ΡI</font></p>
<p>�b���A�S�P�©Ҧ�����ڭ̦���o�i���B�͡A�δ��X�N�������͡I</p>
<p>����H�b�S���@�̪��ѭ��P�N�U�Ф��n���{�ǩM���v�I</p>
<p>�p���ðݥi�V�ӵ{�ǳ]�p�H��<a href="mailto:Seven.s-888@yahoo.com.tw">�����`��</a>�q�l:Seven.s-888@yahoo.com.tw
------------2003�~2��18��--------------</p>
</body>
</html>
