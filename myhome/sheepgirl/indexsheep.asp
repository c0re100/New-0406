<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>

<HTML><HEAD><TITLE>�F�C����-����t��|</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="jh/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=596>
  <TBODY> 
  <TR>
    <TD background=jh/history_top_bg.gif vAlign=top width=596> 
      <TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9 
      width="97%">
        <TBODY> 
        <TR>
          <TD height=25 width="31%">&nbsp;</TD>
          <TD height=25 vAlign=top width="37%"> 
            <DIV align=center><font color="#000000">����t��|</font></DIV>
          </TD>
          <TD height=25 vAlign=top 
width="32%">&nbsp;</TD>
        </TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD background=jh/history_table_bg.gif height=158 vAlign=top width="596"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=596>
        <TBODY> 
        <TR>
          <TD vAlign=top bgcolor="#000000"><BR>
            <TABLE align=center background=jh/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=p9 width="90%">
              <TBODY> 
              <TR>
                <TD align=middle vAlign=center> 
                  <P class=p9><font color="#FF0000"><b>�i���򴣥ܡj</b></font><br>
                    �o��O���򤤳̤j���t��|�A�ѩ�B�b����å@�A�ҥH�t��|�̪��t��ܦh�A�A�i�H�b�o���10�U<br>
                    ��Ȥl�N�i�H��@�өt��^�a�C�A�C�ѳ޾i�L�A���U�L�A�A���Ĥl���j�|�����A���C����@�q�ɶ���A�L�|���u���A�Ȩ��l���A��M�A���F���l�N�i�H�b��ѫǸ̨ϥμɨ��W�[�¤O�B�٥i�H�R�|�O���@�C<br>      
                  </P>
                  <table width="550" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td width="9">&nbsp;</td>
                      <td width="46" valign="top"> 
                        <div align="center"> </div>
                      </td>
                      <td valign="top" width="476"> 
                        <div align="center"> 
                          <div align="center"> 
                            <div align="center"> 
                              <table border="0" width="468" cellspacing="1" cellpadding="0"
            height="20">
                                <center>
                                </center>
                              </table>
                            </div>
                          </div>
                          <div align="center"> 
                            <center>
                              <table border="0" width="386" cellspacing="1" cellpadding="0" class="p9">
                                <tr> 
                                  <td width="100%"> 
                                    <table border="0" width="79%" cellspacing="1" cellpadding="0">
                                      <tr> 
                                        <td class="p3" width="100%">&nbsp;</td>
                                      </tr>
                                      <tr> 
                                        <td class="p2" width="100%"> 
                                          <p align="center"><font color="#FFCC00" class="p9">����t��|�̪��t��:</font><%=10000-tempcount%></p>
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td class="p9" width="100%">��-��i�t��</td>
                                      </tr>
                                    </table>
                                    <table border="0" width="470" cellspacing="1" cellpadding="0">
                                      <tr> 
                                        <td class="p9" width="168" height="20">��-���A���Ĥl���@�Ӧnť���W�r�a�G)</td>
                                        <td class="p3" width="299" rowspan="2"> 
                                          <form name="form1" method="POST" action="buysheep.asp">
                                            <input type="text" name="sheepname" size="33"
                        style="font-family: �s�ө���; font-size: 12px">
                                            <input
                        type="button" value="��i" name="B1"
                        style="font-family: �s�ө���; font-size: 12px">
                                          </form>
                                          <script language="vbscript"> 
<!-- 
sub b1_onclick 
if form1.sheepname.value="" then 
msgbox"�p�Ī��W�r���ର��" 
else 
form1.submit 
end if 
end sub 
--> 
</script>
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td class="p3" width="168" height="20"></td>
                                      </tr>
                                    </table>
                                  </td>
                                </tr>
                              </table>
                            </center>
                          </div>
                          <!--#include file="data1.asp"-->
                          <p><font color="#FFCC00"><span class="p9">��u��F�G</span> 
                            <%rs.open("select * from sheep where user='"&info(0)&"'"),conn,1,1                 
if  rs.bof then%>
                            <span class="p9">�֧־i�Ӥp�Ĭ��A�ȿ��աI</span>       
                            <%else                 
 if rs("feedsheepday")>=3 then%>
                            </font></p>
                          <form method="POST" action="sellmilk.asp">
                            <p><span class="p9">�A�Ĥl���A�b�t��|�̥��F<%=rs("milk")%>�u�@���C�A�Q�V�t��|���h�֤p�ɪ��u���O�H 
                              <input   
          type="text" name="milk" size="9"   
          style="font-family: �s�ө���; font-size: 12px">
                              <input type="submit"   
          value="���u���F" name="B2"   
          style="font-family: �s�ө���; font-size: 12px">
                              <input type="hidden"   
          name="sheepname" value="<%=rs("sheepname")%>">
                              <%else%>
                              �A���Ĥl�٨S���u�@�O</span>�C</p>
                          </form>
                          <%end if%>
                          
                            <%end if%>
                        </div>
                      </td>
                      <td width="97">&nbsp;</td>
                    </tr>
                  </table>
                  <P class=p9>&nbsp; </P>
                  <P align=center class="p9"><a href="javascript:window.close()"><font color="#FFFFFF">�������f</font></a></P>
                  </TD></TR></TBODY></TABLE>
            <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="596"><IMG height=1 src="jh/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>
