<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
action=Request.Querystring("action")
my=info(0)
%>
<HTML>
<HEAD>
<TITLE>���n�ڹC�O</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<style type="text/css">
<!--
td {  font-family: "�s�ө���"; font-size: 12px}
a:link {  font-family: "�s�ө���"; font-size: 12px; color: #000000; text-decoration: none}
a:hover {  font-family: "�s�ө���"; font-size: 12px; color: #FFCCCC; text-decoration: underline}
input {  font-family: "�s�ө���"; font-size: 12px; border: thin #000000 dotted; background-color: #FFFFCC}
-->
</style>
</HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<TABLE border=0 cellPadding=0 cellSpacing=0 width=552 align="center">
  <TBODY> 
  <TR>
    <TD background=../pic/tdbg2.gif width="11">&nbsp;</TD>
    <TD align=middle vAlign=top width=661>
      <TABLE border=0 cellPadding=0 cellSpacing=0 width="601">
        <TBODY>
        <TR>
          <TD width="11"><IMG src="../pic/empty.gif" width="1" height="1"></TD>
          <TD noWrap width="557">
            <TABLE border=0 cellPadding=0 cellSpacing=0>
              <TBODY>
              <TR>
                <TD><IMG src="../pic/icont1.gif" width="30" height="29"></TD>
                <TD>
                  <TABLE border=0 cellPadding=0 cellSpacing=0 height="26" width="518">
                    <TBODY>
                    <TR>
                      <TD rowSpan=3 height="26" width="14">&nbsp;</TD>
                      <TD rowSpan=3 height="26" width="11"><IMG src="../pic/wtdr1.gif" width="11" height="21"></TD>
                      <TD background=../pic/wtdrbg1.gif height="5" width="122"><IMG height=2 
                        src="../pic/empty.gif" width=103></TD>
                      <TD rowSpan=3 height="26" width="19"><IMG src="../pic/wtdr2.gif" width="9" height="21"></TD>
                      <TD rowSpan=3 height="26" width="352" align="right"><a href="javascript:window.close();"><img border="0" src="../pic/close.gif"></a></TD>
                    </TR>
                    <TR>
                      <TD align=middle class=redtext height=17 width="122">::<font size="2" color="#FD794D"><b>���n�ڹC�O</b></font>::</TD>
                    </TR>
                    <TR>
                      <TD background=../pic/wtdrbg2.gif height="6" width="122"><IMG height=4 
                        src="../pic/empty.gif" 
                width=103></TD>
                    </TR></TBODY></TABLE></TD></TR></TBODY></TABLE></TD>
          <TD vAlign=bottom width="28"><IMG src="../pic/rightct12.gif" width="27" height="31"></TD>
          <TD vAlign=bottom width="19"><IMG src="../pic/rightct13.gif" width="19" height="31"></TD></TR>
        <TR>
          <TD width="11"><IMG src="../pic/rightct3.gif" width="11" height="13"></TD>
          <TD background=../pic/rightct4.gif width="557">&nbsp;</TD>
          <TD background=../pic/rightct4.gif width="28">&nbsp;</TD>
          <TD width="19"><IMG src="../pic/rightct14.gif" width="19" height="13"></TD></TR>
        <TR>
          <TD background=../pic/rightct6.gif rowSpan=2 width="11">&nbsp;</TD>
          <TD class=bg1 colSpan=2 rowSpan=2 
          style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" vAlign=top 
            width=557>
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="45%"><img src="water.jpg" width="300" height="150"></td>
                <td width="55%" valign="middle" align="center"><br>
<%
if action="show" then
 Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs=conn.execute("select ��O from �Τ� where id="&info(9))
tl=rs("��O")
if tl<0 then
conn.execute("update �Τ� set ���A='��' where id="&info(9))

	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../exit.asp"
	response.end
Session.Abandon
end if
Randomize timer
r=int(Rnd*10)
Select Case r
 Case 1
     conn.execute("update �Τ� set �Ȩ�=�Ȩ�+80000 where id="&info(9))
     mess="�C���~���ߨ�80000��Ȥl�I�I�I"
 Case 2
     conn.execute("update �Τ� set �D�w=�D�w+10 where id="&info(9))
     mess="<img src='04.gif' width='200' height='100'><br><br>�C���~�����U�A�ҺإСA�D�w�[10�I�I�I"
 Case 3
     conn.execute("update �Τ� set ��O=��O+2000 where id="&info(9))
     mess="�~�g�m�����ȴ̡n��O��_2000�I�I�I"
 Case 4
     conn.execute("update �Τ� set �Ȩ�=�Ȩ�-500 where id="&info(9))
     mess="�C���~�����p�ߥᥢ�F500��Ȥl�I�I�I"
 Case 5
     conn.execute("update �Τ� set �D�w=�D�w+20,�Ȩ�=�Ȩ�+800 where id="&info(9))
     mess="<img src='02.jpg' width='200' height='100'><br><br>�����g�꦳�\�D�w�[20�Ȩ�[800�I�I�I"
 Case 6
     conn.execute("update �Τ� set ��O=��O-500,���O=���O-100 where id="&info(9))
     mess="<img src='02.jpg' width='200' height='100'><br><br>�P�g����q�ɨ��ˡA��O�U��500,���O�U��100�I�I�I"
 Case 7
     conn.execute("update �Τ� set ���O=���O+300 where id="&info(9))
     mess="<img src='01.gif' width='200' height='100'><br><br>���J���H���I�A���O����300�I�I�I"
 Case 8
     conn.execute("update �Τ� set ��O=��O-2000 where id="&info(9))
     mess="<img src='03.jpg' width='200' height='100'><br><br>�j�L�����Ѱ��a�p�A��O�j�l2000��@����Z���I�I�I"
 Case 9
     sex=info(1)
     if sex="�k" then
     conn.execute("update �Τ� set ���O=���O-300,��O=��O-300 where id="&info(9))
     mess="�C���~�����d�ʩ�K�|�s�ӡA���O��O�j��300�I�I�I"
     else
     conn.execute("update �Τ� set ���O=���O-200,��O=��O-200 where id="&info(9))
     mess="�C���~�����J�k��R�������A���O��O�j��200�I�I�I"
     end if
 Case else
     mess="�ȳ~���@���ӥ��L�ơI�I�I"
End Select

	set rs=nothing
	conn.close
	set conn=nothing
response.write mess
%>
<br>
                  <br>
                  <input type="button" name="Button" value="�A���@��S�p��" OnClick="javascript:location.href='index.asp?action=show'">
                  <br>
<%
else
%>
                  �w��<%=my%>�Ӧ��n���m�C���I�I�I<br>
                  10���H���ƥ�I�I�I<br>
                  <input type="button" name="Button" value="�D���@�뤣�i" OnClick="javascript:location.href='index.asp?action=show'">
                  <br>
<% end if %>
                </td>
              </tr>
              <tr> 
                <td colspan="2">&nbsp;</td>
              </tr>
            </table>
          </TD>     
          <TD background=../pic/rightct8.gif vAlign=top width="19"><IMG      
            src="../pic/rightct15.gif" width="19" height="58"></TD></TR>     
        <TR>     
          <TD background=../pic/rightct8.gif width="19">&nbsp;</TD>
        </TR>     
        <TR>     
          <TD width="11"><IMG src="../pic/rightct9.gif" width="11" height="10"></TD>     
          <TD background=../pic/rightct10.gif colSpan=2 width="587">&nbsp;</TD>     
          <TD width="19"><IMG src="../pic/rightct11.gif" width="19" height="10"></TD></TR></TBODY></TABLE>     
    </TD>    
    <TD background=../pic/tdbg2.gif width="13">&nbsp;</TD>
  </TR>    
  <TR>    
    <TD width="11"><IMG src="../pic/tdbgr1.gif" width="10" height="15"></TD>    
    <TD background=../pic/tdbg3.gif width="661"> </TD>    
    <TD width="13"><IMG src="../pic/tdbgr2.gif" width="13" height="15"></TD></TR></TBODY></TABLE>  
    </BODY>
</HTML>