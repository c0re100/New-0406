<%@ LANGUAGE = VBScript %>
<!--#include file="data.asp" -->
<!--#include file="../inc/function.inc.asp" -->
<%
news_type=request("type")
if (news_type="") then
	news_type=1
end if 
%>
<HTML>
<HEAD>
<TITLE></TITLE>
<link href="../../csss.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY leftmargin="0" topmargin="0">
<CENTER>
  <table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF"> 
    <tr> 
      <td colspan="4" height="20" align="center"> 
        <table border="0" width="100%" cellpadding="0" cellspacing="0"> 
          <tr> 
            <td width="160" bgcolor="#284070" align="center" height="20"><font color="#FFFFFF">�F�C����޲z�t��</font></td>
            <td bgcolor="#4880a8" height="20"> 
            </td>
          </tr> 
        </table>
      </td>
    </tr> 
  </table>
  <TABLE border="0" width="778" cellpadding="0" cellspacing="0"> 
    <TR> 
      <TD width="160" valign="top" bgcolor="#4880a8"> 
        <!--#include file="Menu.asp" -->
        <BR>
      </TD>
      <TD valign="top"> 
        <TABLE border="0" width="100%" cellpadding="0" cellspacing="0"> 
          <TR> 
            <TD height="21" bgcolor="#b8d4e8">&nbsp;�޲z �� <%=arr_news_type(news_type,1)%></TD>
          </TR>
          <TR> 
            <TD> 
              <TABLE border="0" width="100%" cellpadding="0" cellspacing="0"> 
                <TR> 
                  <TD align="right" valign="middle">&nbsp;</TD>
                  <TD width="40" height="24">&nbsp;</TD>
                </TR> 
              </TABLE>
            </TD>
          </TR>
          <TR> 
            <TD align="center"> 
                
              <table width="96%" cellpadding="3" cellspacing="1" border="0" bgcolor="#FFFFFF" align="center" class="FormTable">
                <form name="form1" method="post" action="news_act.asp">
                  <tr> 
                    <td align="center" width="100" class="FormTd1">���D�G</td>
                    <td class="FormTd2"> 
                      <input size="40" type="text" name="title">
                    </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">���e�G</td>
                    <td class="FormTd2"> 
                      <textarea rows="8" cols="60" name="content"></textarea>
                    </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">�o�G�ɶ��G</td>
                    <td class="FormTd2"> 
                      <input type="text" name="filltime_year" size=4 maxlength="4" value="<%=year(now)%>">
                      �~ 
                      <%call makeselectmonth("filltime_month",month(now))%>
                      �� 
                      <%call makeselectday("filltime_day",day(now))%>
                      �� </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">�o�G �H�G</td>
                    <td class="FormTd2"> 
                      <input size="40" name="author">
                    </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">�K�@�@�ۡG</td>
                    <td class="FormTd2"> 
                      <input size="40" name="news_form">
                      �i���� </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1" height="30">&nbsp;</td>
                    <td height="30" class="FormTd2" align="center"> 
                      <input class="button" name="Submit2" type="submit" value="�E�e�X�E">
                      <input class="button" name="Submit22" type="reset" value="�E�����E">
                      <input type="button" name="Submit3" value="�E�h�X�E"  onClick="window.history.back()" class="button">
                      <input type="hidden" name="news_type" value="<%=request("type")%>">
                      <input type="hidden" name="action" value="add">
                    </td>
                  </tr>
                </form>
              </table>
            </TD>
          </TR> 
        </TABLE>
      </TD>
    </TR> 
  </TABLE>
</CENTER>
</BODY>
</HTML>
