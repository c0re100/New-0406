<%

' --- �U���B�{���e�]�w   �a�t�@������ ���� http://www.2vh.net
luck1 = "<img src=img/01.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>�j�N</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>���G�A���Ѫ��B��ܩ���!<br>������Q�����ƫK�h���էa,<br>�γ\�|�����G!<br>�S�Ϊ̻��֥h�R�ӱm���a!"
luck2 = "<img src=img/02.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFFF>���N</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>���ߧA�r!<br>�������ӨƨƦp�N,<br>�γ\�a�Ѯ�]�|�സ�O!<br>���L�Z�Ƴ��O�}���a�n�F,<br>�o���O�ؤp���N!"
luck3 = "<img src=img/03.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c00000><font color=#FFFFF>�p�N</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>�߱��]�⤣���F�a!<br>���M���ɷ|�J�쥢��,<br>�����ũ���!"
luck4 = "<img src=img/04.gif border=0></td></tr><tr><td align=center height=15 bgcolor#c00000><font color=#FFFFF>�p��</font></td></tr><tr><td bgcolor=#FFdbdb align=center valign=top height=130>�]�\���Ѥ��A�y�z�j��,<br>�έ�ı�N�n�F!"
luck5 = "<img src=img/05.gif border=0></td></tr><tr><td align=center height=15 bgcolor=#c0000><font color=#FFFFFF>����</font></td></tr><tr><td bgcolor#FFdbdb align=center valign=top height=130>�����n�b�G�o�ǵL�᪱�N!<br>�Ʊ��n�P�a�O�ѧA�M�w��,<br>�O�Ǥߧr!"


randomize
x = int(rnd * 100)
' �T�v 20%
if x < 20 then 
 luck = luck1
' �T�v 20%
elseif x < 40 then
 luck = luck2
' �T�v 30%
elseif x < 70 then
 luck = luck3
' �T�v 20%
elseif x < 90 then
 luck = luck4
' �T�v 10%
else luck = luck5
end if
%>
<html>

<head>
<meta http-equiv="Content-Type"
content="text/html; charset=big5">
<title>����B�{</title>
<STYLE TYPE="text/css">
<!--
tr, td,body,th    {font-size: 9pt}
a:link    {font-size: 9pt; text-decoration:none; color:<%=link%> }
a:visited {font-size: 9pt; text-decoration:none; color:<%=vlink%> }
a:active  {font-size: 9pt; text-decoration:none; color:<%=alink%> }
a:hover   {font-size: 9pt; text-decoration:underline; color:<%=alink%> }
input,select,textarea     {font-size: 9pt; border: 1 solid black}
-->
</STYLE>
</head>

<body bgcolor="#FFFFFF" topmargin="0" leftmargin="0">
<div align="left">

<table border="1" cellspacing="1" width="157"
bordercolor="#FFADAD">
    <tr>
        <td height=142 align=center valign=middle><%=luck%>
        </td>
    </tr>
</table>
</div>
</body>
</html>
