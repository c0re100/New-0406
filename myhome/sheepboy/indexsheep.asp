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
rs.open "select 道德,配偶 from 用戶 where id="&info(9),conn
daode=int(rs("道德"))
if daode>5000 and rs("配偶")="無" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
 Response.Write "<script language=JavaScript>{alert('你結婚了嗎？如果是那你的道德實在太低了，多為自己將來的小孩積點德吧！');window.close();}</script>"
Response.End
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>

<HTML><HEAD><TITLE>靈劍江湖-江湖育兒院</TITLE>
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
            <DIV align=center><font color="#000000">江湖育兒院</font></DIV>
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
                      江湖中的夫婦可以來這裡檢查身體生小孩，但夫妻倆的道德必須在5000點以上，嘿嘿，壞人也想要小孩嗎，去托兒所領一個吧。生小孩的收費為100萬銀兩，沒錢的你就去托兒所領一個吧。有了小孩後，夫婦倆都可以來照顧，當然，江湖事非多，我們會幫你帶養你的孩子，不過也是常回來看看哦。你的孩子長大了可以幫你賺錢，也可以幫你打架，嘿嘿，本程序來自江湖中人安德森大俠的提示，程序在測試中，希望大家發現問題及時與站長聯繫。<br>      
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
                                          <p align="center"><font color="#FFCC00" class="p9">江湖育兒院</font></p>
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
                                          <p>□-請你寫上你妻子（丈夫）的名字，再為你們所生的孩子取個好聽一點的名字吧：)</p>
                                          </td>
                                        <td class="p3" width="292"> 
                                          <form name="form1" method="POST" action="buysheep.asp">
                                            <input type="text" name="sheepname2" size="8"
                        style="font-family: 新細明體; font-size: 12px">
                                            <input type="text" name="sheepname" size="24"
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
                          <%end if
rs.close
	set rs=nothing
	%>
                          <a href="feedsheep.asp"><font color="#FFFFFF"><br>
                          <br>
                          照顧小孩</font></a> </div>
                      </td>
                      <td width="10">&nbsp;</td>
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
