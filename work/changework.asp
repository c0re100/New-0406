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
<html>
<link rel="stylesheet" href="../css.css">
<title>¾�~���Щ�</title>
<body bgcolor="#996699" background="../chat/bk.jpg" leftmargin="0" topmargin="0">
<table border="0" cellspacing="0" cellpadding="0" width="382" align="center">
  <tr> <td height="81" valign="top">
      <form method=POST action='cwork.asp'> 
        <table width="366" BORDER="1" align="center" background="../ff.jpg">
          <tr> <td> <tr> <td align=center>�п�ܱz��¾�~�G 
<select name=jiu size=1>
                <option value="�Ԥh" selected>�Ԥh</option>
                <option value="�ɲ��Ԥh">�ɲ��Ԥh</option>
                <option value="���Z�Ԥh">���Z�Ԥh</option>
                <option value="���ҾԤh">���ҾԤh</option>
                <option value="���s�Ԥh">���s�Ԥh</option>
                <option value="�i�h">�i�h</option>
                <option value="�ʾԫi�h">�ʾԫi�h</option>
                <option value="��«i�h">��«i�h</option>
                <option value="���s�i�h">���s�i�h</option>
                <option value="����i�h">����i�h</option>
                <option value="�D�h">�D�h</option>
                <option value="���D�h">���D�h</option>
                <option value="���k�v">���k�v</option>
                <option value="���u�H">���u�H</option>
                <option value="���Ѯv">���Ѯv</option>
              </select> </td></tr> <tr> <td  align=center> 
<input type=submit value=�ڿ�ܦn�F name="submit"> </td></tr> <tr> 
            <td valign="top" > 
              <p><FONT SIZE="-1" COLOR="#000000">¾�~�O�z�b����n���g�٨ӷ��A��ܤ��P��¾�~��z���u�@�|�����P���v�T�G<br>
                �Ԥh�G 
                ���ŻݨD0�ťH�W<br>
                <font color="#FF0000">�ɲ��Ԥh�G���ŻݨD50�ťH�W ���ǧ�����2�� <br>
                ���Z�Ԥh�G���ŻݨD250�ťH�W �y�P5�� ���ǧ�����3�� <BR>
                ���ҾԤh�G���ŻݨD400�ťH�W �y�P�\5�� ���ǧ�����4�� <br>
                ���s�Ԥh�G���ŻݨD650�ťH�W �y�P30�� ���ǧ�����5�� </font> </FONT></p>
              <p><font color="#000000" size="-1">�i�h�G���ŻݨD0�ťH�W<br>
                <font color="#000099">�ʾԫi�h�G���ŻݨD80�ťH�W ���Ǩ��s��2�� <br>
                ��«i�h�G���ŻݨD280�ťH�W �y�P5�� ���Ǩ��s��3�� <br>
                ���s�i�h�G���ŻݨD500�ťH�W �y�P�\5�� ���Ǩ��s��4�� <br>
                ����i�h�G���ŻݨD750�ťH�W �y�P45�� ���Ǩ��s��5�� </font> </font></p>
              <p><font color="#0033FF" size="-1">�D�h�G���ŻݨD0�ťH�W<br>
                ���D�h�G���ŻݨD90�ťH�W 
                ���ǪZ�\��1.5��<br>
                ���k�v�G���ŻݨD320�ťH�W 
                �y�P5�� 
                ���ǪZ�\��2��<br>
                ���u�H�G���ŻݨD550�ťH�W 
                �y�P�\5�� 
                ���ǪZ�\��2.5�� 
                <br>
                ���Ѯv�G���ŻݨD780�ťH�W 
                �y�P65�� 
                ���ǪZ�\��3��<br>
                �D�h�����ɦ�B�Z�\���\��</font></p>
              <p align="center"><font color="#000000">��¾�t����100�U��¾�O</font></p>
            </td></tr> 
</table></form></td></tr> </table>
<div align="center"><font color="#00FF66"><b></b></font> </div>
</body>
</html>
 