<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")%>
<HTML><HEAD><TITLE>�F�C����-���c</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="../pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="bg.jpg">
<div align="center">
  <table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="32%" height="31">
    <tr align="center"> 
      <td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">���c</font></b></td>
    </tr>
  </table>
  <br>
</div>
<table border=0 cellpadding=0 cellspacing=0 width="530" align="center">
  <tbody> 
  <tr> 
    <td class=bg1 colspan=2 rowspan=2 
          style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
      <table border="0" cellpadding="0" cellspacing="0" width="554" height="231">
        <tr> 
          <td align="center" colspan="2"> <font size="2" color="#65251C"><b> <font color="#FFFFFF"><%=info(0)%><font color="#00FF00">�i�J���c�A�b�o��A�A�i�H���ūӪ�������Ѯ@</font></font></b></font> 
          </td>
        </tr>
        <tr> 
          <td align="center" rowspan="2" width="200"><img src="girl.jpg" width="200" height="151"> 
          </td>
          <td width="343"> 
            <p>&nbsp;</p>
            <p><font size="3" color="#65251C"><b><font color="#FFFFFF">���j�Y���D�G</font></b></font><font size="2" color="#00FFFF">�j���A�֨�ڭ̪��c�ӡA�M�ڭ̳o�W�Ū��Өk�ͽͤߡA�A�N��W�[���O�@�A���L�ڭ̳o�઺�Өk���O�������樭�@�A�p�G�A�O�ӭ��A�S�����ܡA�A�]�i�H��ڭ̳o�Ӥu�@�@�A���L�p�G�A�������������ܴN���ӸդF�A�����A�ڭ̳o���n�ӭ��@�C</font> 
            </p>
          </td><br>
        </tr>
        <tr>
          <td width="343" valign="middle" align="center" height="115"> 
            <p>&nbsp;</p>
            <p><font size="2" color="#65251C"><b><a href="xiaojie.asp"><font color="#0000FF">�H�b����A�H�i�v�y�A�Ķi���c</font></a></b></font> 
            </p>
            <p><font size="2" color="#65251C"><b><a href="dengji.asp"><font color="#FF00FF">�u�A�S���F�A���b�o�W�Z��F</font></a><br>
              </b></font></p>
            <p><b><font size="2"><a href="delgirl.asp"><font color="#FFFF00">���z�F�A�o�O�H�L����l�ܡH�ڭnū��</font></a></font></b></p>
            </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan=2>&nbsp;</td>
  </tr>
  </tbody> 
</table>
</BODY></HTML>       
