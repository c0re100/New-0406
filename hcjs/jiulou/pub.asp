<%
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>

<html>
<link rel="stylesheet" href="../../css.css">
<body leftmargin="0" topmargin="0" bgcolor="#660000" text="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="0" width="597">
  <tr> 
    <td bgcolor="#660000" align="right" colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#660000" align="right" width="330" height="328" valign="top"> 
      <p align=center style='color:yellow'><font color="#FF9900">�w��<%=info(0)%>���{�K�P��</font> 
        <br>
      <table border="0" align="center" cellpadding="0" cellspacing="0" width="300" height="200">
        <tr> 
          <td colspan="3"><font color="#FF9900">�o�̬O�ʫ����W�P�s��,�����̭��ټM�n���_�n�����x.�b�s�Ӫ��f���@�ӵP�l�W���g��:</font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> �s �W </font></td>
          <td><font color="#FF9900"> �� �O </font></td>
          <td><font color="#FF9900"> �� �� </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> �Ѧʤz </font></td>
          <td><font color="#FF9900"> 50 </font></td>
          <td><font color="#FF9900"> 250 </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> ��³�G </font></td>
          <td><font color="#FF9900"> 60 </font></td>
          <td><font color="#FF9900"> 300 </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> ���d </font></td>
          <td><font color="#FF9900"> 70 </font></td>
          <td><font color="#FF9900"> 350 </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> �T�O </font></td>
          <td><font color="#FF9900"> 80 </font></td>
          <td><font color="#FF9900"> 400 </font></td>
        </tr>
      </table>
    </td>
    <td align="left" height="328" width="59">&nbsp;</td>
    <td align="center" width=300 height="328" valign="top"> <br>
      <form method=POST action='pub1.asp'>
        <table width="300">
          <tr> 
            <td> 
          <tr> 
            <td align=center>�A�V�p�G�R�U�F�G 
              <select name=jiu size=1>
                <option value="lao">�Ѧʤz 
                <option value="wu">��³�G 
                <option value="du">���d 
                <option value="mao">�T�O 
              </select>
              �N�|�_�s�M</td>
          </tr>
          <tr> 
            <td  align=center> 
              <input type=submit value=�@���Ӻ� name="submit">
            </td>
          </tr>
          <tr> 
            <td > 1�B�b�s�ӳܰs�i�H�W�[��O�F <br>
              2�B�s���M�O�Ӧn�F��,���O����g�h�F �s�q�����Фֶ�,�_�h,�H�H......<br>
              3�B�s�q�P�A���Z�\�M�g�禳���I �s�q=(�Z�\+�g��)/100 �n�Q���K�˶� �s�q>1�A �ۤv�i�q�i�q�a�I</td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
</body>
</html>


