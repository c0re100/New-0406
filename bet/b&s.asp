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
rs.open "Select �Ȩ� from �Τ� where id="&info(9),conn
yin=rs("�Ȩ�")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML>
<HEAD>
<title>��j�p</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<body text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../images/8.jpg">
<div align="left"></div>
<div align=center> 
  <p><font color="#000000" size="+3">��j�p</font><font size="2"><br>
    �̤֤U�`�Ȥl�O <b><font color="#CC0000">10 </font> ��<br>
    </b>�̤j�U�`�Ȥl�O <font color="#CC0000"><b>3000</b></font><b> ��</b> <br>
    �A�{�b�@�@���Ȥl <b><font color="#CC0000"><%=yin%></font> ��</b> �i�H�@����`</font></p>
  <table width="545" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td> 
        <form method="POST" action="b&amp;spose.asp">
          <table border=1 cellspacing=0 cellpadding=0 align="center" width="350" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr align="center" bgcolor="#FFFFFF"> 
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
            </tr>
            <tr bgcolor="#336633"> 
              <td width="960" colspan="3" height="29"><font size="2" class="c" color="#000000"><b>&nbsp;&nbsp;<font color="#FFFFFF">�п��</font></b></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td align=center colspan="3"> 
                <table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                  <tr align="center"> 
                    <td width="50%"><img src="../jhimg/bbs/big.gif" width="46" height="40"></td>
                    <td width="50%"><img src="../jhimg/bbs/small.gif" width="46" height="40"></td>
                  </tr>
                  <tr align="center"> 
                    <td width="50%"> 
                      <input type="radio" name="select" value="big" checked>
                    </td>
                    <td width="50%"> 
                      <input type="radio" name="select" value="small">
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr> 
              <td align=center colspan="3"><font size="2" color="#000000">�ڭn�U�`�G 
                <input type="text" name="splosh" size="4" value="0" maxlength="4">
                &nbsp;<b>��</b></font></td>
            </tr>
            <tr> 
              <td align=center colspan="3"> <font size="2" color="#000000"> 
                <input type="submit" value="�U�`�աI�I�I" name="B1" >
                <input type="reset" value="�ڭn�Ҽ{�@�U�G�^" name="B2" >
                </font></td>
            </tr>
          </table>
        </form>
        <p align="center"><font size="2">ĵ�i�G�C������A�A���y�O�ȷ|�� 1 �I</font></p>
      </td>
    </tr>
  </table>
<p align="center">�F�C�����v�Ҧ�   </p>
</div>
</BODY>
</HTML>
