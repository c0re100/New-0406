<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
nickname=info(0)

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �D�w,�t�� from �Τ� where id="&info(9),conn
daode=int(rs("�D�w"))
if daode>5000 and rs("�t��")="�L" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 Response.Write "<script language=JavaScript>{alert('�A���B�F�ܡH�p�G�O���A���D�w��b�ӧC�F�A�h���ۤv�N�Ӫ��p�Ŀn�I�w�a�I');window.close();}</script>"
Response.End
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>

<HTML><HEAD><TITLE>�F�C����-����|��|</TITLE>
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
            <DIV align=center><font color="#000000">����|��|</font></DIV>
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
                      ���򤤪��Ұ��i�H�ӳo���ˬd����ͤp�ġA���ҩd�Ǫ��D�w�����b5000�I�H�W�A�K�K�A�a�H�]�Q�n�p�ĶܡA�h����һ�@�ӧa�C�ͤp�Ī����O��100�U�Ȩ�A�S�����A�N�h����һ�@�ӧa�C���F�p�ī�A�Ұ��ǳ��i�H�ӷ��U�A��M�A����ƫD�h�A�ڭ̷|���A�a�i�A���Ĥl�A���L�]�O�`�^�Ӭݬݮ@�C�A���Ĥl���j�F�i�H���A�ȿ��A�]�i�H���A���[�A�K�K�A���{�ǨӦۦ��򤤤H�w�w�ˤj�L�����ܡA�{�Ǧb���դ��A�Ʊ�j�a�o�{���D�ήɻP�����pô�C<br>      
                  </P>
                  <table width="550" border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr> 
                      <td width="4">&nbsp;</td>
                      <td width="23" valign="top"> 
                        <div align="center"> </div>
                      </td>
                      <td valign="top" width="522"> 
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
                              <table border="0" width="493" cellspacing="1" cellpadding="0" class="p9">
                                <tr> 
                                  <td width="100%"> 
                                    <table border="0" width="100%" cellspacing="1" cellpadding="0">
                                      <tr> 
                                        <td class="p3" width="100%" colspan="2">&nbsp;</td>
                                      </tr>
                                      <tr> 
                                        <td class="p2" width="100%" colspan="2"> 
                                          <p align="center"><font color="#FFCC00" class="p9">����|��|</font></p>
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td class="p9" width="50%" height="26">&nbsp;</td>
                                        <td class="p9" width="50%" height="26">&nbsp; 
                                        </td>
                                      </tr>
                                    </table>
                                    <table border="0" width="493" cellspacing="1" cellpadding="0">
                                      <tr> 
                                        <td class="p9" height="41">
                                          <p>��-�ЧA�g�W�A�d�l�]�V�ҡ^���W�r�A�A���A�̩ҥͪ��Ĥl���Ӧnť�@�I���W�r�a�G)</p>
                                          </td>
                                        <td class="p3" width="292"> 
                                          <form name="form1" method="POST" action="buysheep.asp">
                                            <input type="text" name="sheepname2" size="8"
                        style="font-family: �s�ө���; font-size: 12px">
                                            <input type="text" name="sheepname" size="24"
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
                          <%end if
rs.close
	set rs=nothing
	%>
                          <a href="feedsheep.asp"><font color="#FFFFFF"><br>
                          <br>
                          ���U�p��</font></a> </div>
                      </td>
                      <td width="10">&nbsp;</td>
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
