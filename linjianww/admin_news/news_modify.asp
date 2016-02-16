<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../../chat/readonly/bomb.htm"
if info(2)<10   then Response.Redirect "../error.asp?id=900"%>
<!--#include file="data.asp" -->
<%
news_type=request("type")
if (news_type="") then
	news_type=1
end if 
%>
<HTML>
<HEAD>
<TITLE>艶糃打恨瞶╰参</TITLE>
<link href="../../csss.css" rel="stylesheet" type="text/css">
</HEAD>
<BODY leftmargin="0" topmargin="0">
<CENTER>
  <table width="778" border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#FFFFFF"> 
    <tr> 
      <td colspan="4" height="20" align="center"> 
        <table border="0" width="100%" cellpadding="0" cellspacing="0"> 
          <tr> 
            <td width="160" bgcolor="#284070" align="center" height="20"><font color="#FFFFFF">艶糃打恨瞶╰参</font></td>
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
      </TD>
      <TD valign="top">
      <TABLE border="0" width="100%" cellpadding="0" cellspacing="0">
          <TR>
            <TD height="21" bgcolor="#b8d4e8">&nbsp;ず甧э</TD>
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
              <!--#include file="../inc/function.inc.asp" -->
              <form name="form1" method="post" action="news_act.asp">
                <table width="96%" cellpadding="2" cellspacing="1" border="0" class="FormTable">
<%
	id=request("id")
	sql="select * from hc_news where id="&id
	Set rs=Server.CreateObject("ADODB.Recordset")
	rs.Open sql,Conn,adOpenStatic, adLockReadOnly, adCmdText
	if NOT rs.EOF then 
%>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">夹肈</td>
                    <td class="FormTd2"> 
                      <input size="60" type="text" name="title" value="<%=rs("title")%>">
                    </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">ず甧</td>
                    <td class="FormTd2"> 
                      <textarea rows="8" cols="60" name="content"><%=rs("content")%></textarea>
                    </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">祇丁</td>
                    <td class="FormTd2">
                      <input type="text" name="filltime_year" size=4 maxlength="4" value="<%=year(rs("filltime"))%>">
                       
                      <%call makeselectmonth("filltime_month",month(rs("filltime")))%>
                      る 
                      <%call makeselectday("filltime_day",day(rs("filltime")))%>
                      ら </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1"></td>
                    <td class="FormTd2"> 
                      <input size="20" type="text" name="author" value="<%=rs("author")%>">
                    </td>
                  </tr>
                  <tr> 
                    <td align="center" width="100" class="FormTd1">篕</td>
                    <td class="FormTd2"> 
                      <input size="20" type="text" name="news_form" value="<%=rs("news_form")%>">
                    </td>
                  </tr>
                  <tr> 
                    <td align="right" width="100" height="30" class="FormTd1">&nbsp;</td>
                    <td height="30" class="FormTd2" align="center"> 
                      <input class="button" name="Submit2" type="submit" value="э">
                      <input class="button" name="Submit2" type="button" value="癶" onClick="JavaScript:window.history.back()">
                      <input type="hidden" name="action" value="modify">
                      <input type="hidden" name="id" value="<%=request("id")%>">
                      <input type="hidden" name="news_type" value="<%=rs("type")%>">
                    </td>
                  </tr>
                </table>
              </form>
            </TD>
          </TR>
            <%
	else
		echo "⊿Τ计沮"
	end if 
	rs.close
%>
      </TABLE>
      </TD>
    </TR>
</TABLE>
</CENTER>
</BODY>
</HTML>
