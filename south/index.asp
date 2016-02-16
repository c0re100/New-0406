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
<TITLE>江南夢遊記</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type>
<style type="text/css">
<!--
td {  font-family: "新細明體"; font-size: 12px}
a:link {  font-family: "新細明體"; font-size: 12px; color: #000000; text-decoration: none}
a:hover {  font-family: "新細明體"; font-size: 12px; color: #FFCCCC; text-decoration: underline}
input {  font-family: "新細明體"; font-size: 12px; border: thin #000000 dotted; background-color: #FFFFCC}
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
                      <TD align=middle class=redtext height=17 width="122">::<font size="2" color="#FD794D"><b>江南夢遊記</b></font>::</TD>
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
rs=conn.execute("select 體力 from 用戶 where id="&info(9))
tl=rs("體力")
if tl<0 then
conn.execute("update 用戶 set 狀態='死' where id="&info(9))

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
     conn.execute("update 用戶 set 銀兩=銀兩+80000 where id="&info(9))
     mess="遊玩途中撿到80000兩銀子！！！"
 Case 2
     conn.execute("update 用戶 set 道德=道德+10 where id="&info(9))
     mess="<img src='04.gif' width='200' height='100'><br><br>遊玩途中幫助農夫種田，道德加10！！！"
 Case 3
     conn.execute("update 用戶 set 體力=體力+2000 where id="&info(9))
     mess="途經《有間客棧》體力恢復2000！！！"
 Case 4
     conn.execute("update 用戶 set 銀兩=銀兩-500 where id="&info(9))
     mess="遊玩途中不小心丟失了500兩銀子！！！"
 Case 5
     conn.execute("update 用戶 set 道德=道德+20,銀兩=銀兩+800 where id="&info(9))
     mess="<img src='02.jpg' width='200' height='100'><br><br>平息土匪有功道德加20銀兩加800！！！"
 Case 6
     conn.execute("update 用戶 set 體力=體力-500,內力=內力-100 where id="&info(9))
     mess="<img src='02.jpg' width='200' height='100'><br><br>與土匪較量時受傷，體力下降500,內力下降100！！！"
 Case 7
     conn.execute("update 用戶 set 內力=內力+300 where id="&info(9))
     mess="<img src='01.gif' width='200' height='100'><br><br>路遇高人指點，內力提升300！！！"
 Case 8
     conn.execute("update 用戶 set 體力=體力-2000 where id="&info(9))
     mess="<img src='03.jpg' width='200' height='100'><br><br>大俠不知天高地厚，體力大損2000於一次比武中！！！"
 Case 9
     sex=info(1)
     if sex="男" then
     conn.execute("update 用戶 set 內力=內力-300,體力=體力-300 where id="&info(9))
     mess="遊玩途中整日留戀於春院酒樓，內力體力大減300！！！"
     else
     conn.execute("update 用戶 set 內力=內力-200,體力=體力-200 where id="&info(9))
     mess="遊玩途中陷入糊塗愛情之中，內力體力大減200！！！"
     end if
 Case else
     mess="旅途中一路太平無事！！！"
End Select

	set rs=nothing
	conn.close
	set conn=nothing
response.write mess
%>
<br>
                  <br>
                  <input type="button" name="Button" value="再玩一趟又如何" OnClick="javascript:location.href='index.asp?action=show'">
                  <br>
<%
else
%>
                  歡迎<%=my%>來江南水鄉遊玩！！！<br>
                  10個隨機事件！！！<br>
                  <input type="button" name="Button" value="非走一趟不可" OnClick="javascript:location.href='index.asp?action=show'">
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