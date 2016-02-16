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
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<!--#include file="data1.asp"-->

<HTML><HEAD><TITLE>靈劍江湖-江湖育兒所</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="jh/clubcom.css" rel=stylesheet type=text/css>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#000000 text=#ffffff topMargin=0>
<TABLE align=center border=0 cellPadding=0 cellSpacing=0 width=80%>
  <TBODY> 
  <TR>
    <TD background=jh/history_top_bg.gif vAlign=top width=100%> 
      <TABLE align=center border=0 cellPadding=2 cellSpacing=0 class=p9 
      width="100%">
        <TBODY> 
        <TR>
          <TD height=25 width="31%">&nbsp;</TD>
          <TD height=25 vAlign=top width="33%">
            <DIV align=center><font color="#000000">江湖育所院</font></DIV>
          </TD>
          <TD height=25 vAlign=top 
width="36%">&nbsp;</TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD background=jh/history_table_bg.gif height=158 vAlign=top width="100%"> 
      <TABLE align=center border=0 cellPadding=0 cellSpacing=0 class=mountain 
      width=85%>
        <TBODY> 
        <TR>
          <TD vAlign=top width="100%"><BR>
            <TABLE align=center background=jh/table_bg.gif border=0 
            cellPadding=0 cellSpacing=0 class=table1 width="85%">
              <TBODY> 
              <TR>
                <TD align=middle vAlign=center> 
                  <P class=p9><font color="#FF0000"><b>【江湖提示】</b></font><br>
                      這兒是喂養小孩的地方，你們小孩就住在這兒，看著她無憂無慮的生活著，你覺得在江湖裡不在孤獨。不過你們夫妻每天得來這照顧她哦，不然她會死掉的哦，孩子長大後可以幫你們賺錢打架什麼的哦<br>
                  </P>
                  <table border="0" width="476" cellspacing="1" cellpadding="0">
                    <tr> 
                      <td width="472">&nbsp;</td>
                    </tr>
                    <tr> 
                      <td width="472"> 
                        <table border="0" bordercolor="#ffffff" cellpadding="0"
                cellspacing="1" width="470">
                          <tr> 
                            <td> 
                              <form method="POST" action="checksheep.asp">
                                <table border="0" width="100%" cellspacing="1"
                        cellpadding="0" height="89">
                                  <tr> 
                                    <td class="p2" width="100%" height="25"> 
                                      <%
rs.open"select sheepname from sheep where user='"&info(0)&"' or sheep002='"&info(0)&"'",conn,1,1
if rs.bof then
response.write "你們還沒有小孩呢！<a href='indexsheep.asp'><font color='#FFFFFF'>關閉窗口</font>快生一個吧</a>"
else
                              %>
                                      <p align="center"><font color="#FFCC00"><span class="p9">你的孩子的名字:</span><%=rs("sheepname")%> 
                                        <input 
                              type="hidden" name="nick" size="20"
                              style="font-family: Tahoma; font-size: 12px"
                              value="<%=rs("sheepname")%>">
                                        </font>
                                    </td>
                                  </tr>
                                  <tr> 
                                    <td class="p3" width="100%" height="27"> 
                                      <p align="center">
                                        <input type="submit"
                              value="進入" name="B1"
                              style="font-family: 新細明體; font-size: 12px">
                                    </td>
                                  </tr>
                                </table>
                              </form>
                              <% 

end if 
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
                      %>
                            </td>
                          </tr>
                        </table>
                      </td>
                    </tr>
                  </table>
                  <P class=p9>&nbsp; </P>
                  <P align=center class="p9"><b><a href="javascript:window.close()"><font color="#FFFFFF">關閉窗口</font></a><a href="../../jh.asp"></a></b></P>
                  <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE>
            <P align=center>&nbsp;</P></TD></TR></TBODY></TABLE></TD></TR>
  <TR>
    <TD bgColor=#847939 height=1 width="100%"><IMG height=1 src="jh/page_point.gif" 
      width=1></TD>
  </TR></TBODY></TABLE></BODY></HTML>
