<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>�Ŭu�D����</title>
<link rel="StyleSheet" href="../../css.css" title="Contemporary">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>

<body background="../../ff.jpg" leftmargin="0" topmargin="0">
<div align="center">
<center>
    <table width="348" border="0" cellpadding="0" cellspacing="0" bgcolor="#003366">
      <tr bgcolor="#FFFFFF"> 
        <td></td>
      </tr>
      <tr> 
        <td> <form method="POST" action="CHECKSEX.ASP">
            <br>
            <br>
            <table width="100%" cellspacing="1">
              <tr> 
                <td class="p2" width="100%"> <font color="#CCCCCC">�w �� �i �J �F �C 
                  �� �� �s �u �~ �D �� ��&nbsp;</font> </td>
              </tr>
              <tr > 
                <td class="p3" width="100%" height="29"> <font color="#FFFFFF">�i�J�~�D��N�����u��I�I��O1000 
                  </font> </td>
              </tr>
              <tr> 
                <td class="p2" width="100%" height="50"> <input type="radio" value="�k"
name="SEX" checked> <font color="#FFFFFF"> <font color="#CCCCCC">�k����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
                  <input type="radio" value="�k" name="SEX">
                  �k����</font></font> </td>
              </tr>
              <tr> 
                <td class="p3" width="100%"> <p>&nbsp;</p>
                  <p> 
                    <input type="submit" value="�i�J�~�D"
name="B1">
                  </p></td>
              </tr>
            </table>
          </form></td>
      </tr>
    </table>
</center>
</div>

</body>

</html>
