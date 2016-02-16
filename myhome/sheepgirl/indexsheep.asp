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

<HTML><HEAD><TITLE>靈劍江湖-江湖孤兒院</TITLE>
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
            <DIV align=center><font color="#000000">江湖孤兒院</font></DIV>
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
                  <P class=p9><font color="#FF0000"><b>【江湖提示】</b></font><br>
                    這兒是江湖中最大的孤兒院，由於處在江湖亂世，所以孤兒院裡的孤兒很多，你可以在這兒花10萬<br>
                    兩銀子就可以領一個孤兒回家。你每天喂養他，照顧他，你的孩子長大會報答你的。等到一段時間後，他會打工給你賺豆子的，當然你有了豆子就可以在聊天室裡使用暴豆增加威力、還可以買會費的哦。<br>      
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
                                          <p align="center"><font color="#FFCC00" class="p9">江湖孤兒院裡的孤兒:</font><%=10000-tempcount%></p>
                                        </td>
                                      </tr>
                                      <tr> 
                                        <td class="p9" width="100%">□-領養孤兒</td>
                                      </tr>
                                    </table>
                                    <table border="0" width="470" cellspacing="1" cellpadding="0">
                                      <tr> 
                                        <td class="p9" width="168" height="20">□-為你的孩子取一個好聽的名字吧：)</td>
                                        <td class="p3" width="299" rowspan="2"> 
                                          <form name="form1" method="POST" action="buysheep.asp">
                                            <input type="text" name="sheepname" size="33"
                        style="font-family: 新細明體; font-size: 12px">
                                            <input
                        type="button" value="領養" name="B1"
                        style="font-family: 新細明體; font-size: 12px">
                                          </form>
                                          <script language="vbscript"> 
<!-- 
sub b1_onclick 
if form1.sheepname.value="" then 
msgbox"小孩的名字不能為空" 
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
                          <p><font color="#FFCC00"><span class="p9">算工資了：</span> 
                            <%rs.open("select * from sheep where user='"&info(0)&"'"),conn,1,1                 
if  rs.bof then%>
                            <span class="p9">快快養個小孩為你賺錢啦！</span>       
                            <%else                 
 if rs("feedsheepday")>=3 then%>
                            </font></p>
                          <form method="POST" action="sellmilk.asp">
                            <p><span class="p9">你孩子為你在孤兒院裡打了<%=rs("milk")%>工作單位。你想向孤兒院收多少小時的工錢呢？ 
                              <input   
          type="text" name="milk" size="9"   
          style="font-family: 新細明體; font-size: 12px">
                              <input type="submit"   
          value="收工錢了" name="B2"   
          style="font-family: 新細明體; font-size: 12px">
                              <input type="hidden"   
          name="sheepname" value="<%=rs("sheepname")%>">
                              <%else%>
                              你的孩子還沒找到工作呢</span>。</p>
                          </form>
                          <%end if%>
                          
                            <%end if%>
                        </div>
                      </td>
                      <td width="97">&nbsp;</td>
                    </tr>
                  </table>
                  <P class=p9>&nbsp; </P>
                  <P align=center class="p9"><a href="javascript:window.close()"><font color="#FFFFFF">關閉窗口</font></a></P>
                  </TD></TR></TBODY></TABLE>
            <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="596"><IMG height=1 src="jh/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>
