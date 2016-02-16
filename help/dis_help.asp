<%@ LANGUAGE = VBScript %>
<!--#include file="data.asp" -->
<!--#include file="../inc/global.inc.asp" -->
<!--#include file="../inc/function.inc.asp" -->

<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"


id=request("id")
Set rs=Server.CreateObject("ADODB.Recordset")
sql="select * from hc_news where id="& id
rs.Open sql,Conn,1,1
if rs.eof or rs.bof then
       Response.Write "<script Language=Javascript>alert('沒有這個幫助，或者已經被刪!');window.close();</script>"
       rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
       response.end
else

%>

<HTML>
<HEAD>
<TITLE>港龍江湖幫助中心</TITLE>
<STYLE type="text/css">
<!--
TD{
  font-size : 12px;
  font-weight : lighter;
  line-height : 16px;
  color : #643706;
  margin-top : 1px;
  margin-left : 1px;
  margin-right : 1px;
  margin-bottom : 1px;
}
A{
  color : #553006;
  text-decoration : none;
}
-->
</STYLE>
</HEAD>
<BODY topmargin="0" leftmargin="0" rightmargin="0" bgcolor="#e8e3d0">
<CENTER>
<TABLE border="0" width="550" cellpadding="0" cellspacing="0">
  <TBODY>
    <TR>
      <TD><IMG src="../pic/bz_pic.gif" width="550" height="25" border="0"></TD>
    </TR>
    <TR>
      <TD background="../pic/a_image41.gif" align="center">
      <TABLE border="0" cellpadding="0" cellspacing="0">
        <TBODY>
          <TR>
            <TD height="5"></TD>
          </TR>
        </TBODY>
      </TABLE>
      <TABLE width="97%" border="0" cellpadding="2" cellspacing="0">
        <TBODY>
          <TR>
            <TD bgcolor="#e8e3d0" align="center">
            <TABLE width="100%" border="0" cellpadding="0" cellspacing="1">
              <TBODY>
                <TR>
                  <TD colspan="3" nowrap>
                    <table width="98%" border="0" align="center" cellpadding="3">
                      <tr>
                        <td width="65" height="17"><img src="../jh_pic/image4.gif" width="9" height="9"> 
                          <b><a href="help.asp?type=6&name=%C8%EB%C3%C5%C6%AA">入門篇</a></b> 
                        </td>
                        <td width="65" height="17"><img src="../jh_pic/image4.gif" width="9" height="9"> 
                          <b><a href="help.asp?type=7&name=%B7%A2%B2%C6%C6%AA">發財篇</a></b></td>
                        <td width="65" height="17"><img src="../jh_pic/image4.gif" width="9" height="9"> 
                          <b><a href="help.asp?type=8&name=%B3%C9%BC%D2%C6%AA">成家篇</a></b> 
                        </td>
                        <td width="65" height="17"> <img src="../jh_pic/image4.gif" width="9" height="9"> 
                          <b><a href="help.asp?type=9&name=%C1%B7%CE%E4%C6%AA">練武篇</a></b> 
                        </td>
                        <td width="65" height="17"><img src="../jh_pic/image4.gif" width="9" height="9"> 
                          <b><a href="help.asp?type=10&name=%C1%C4%CC%EC%C6%AA">聊天篇</a></b> 
                        </td>
                        <td width="65" height="17"><img src="../jh_pic/image4.gif" width="9" height="9"> 
                          <b><a href="help.asp?type=11&name=%B6%AF%CE%E4%C6%AA">動武篇</a></b> 
                        </td>
                        <td width="68" height="17"><img src="../jh_pic/image4.gif" width="9" height="9"> 
                          <b><a href="help.asp?type=12&name=%B9%DC%C0%ED%C6%AA">管理篇</a></b> 
                        </td>
                      </tr>
                    </table>
                  </TD>
                </TR>
                <TR>
                  <TD width="3">&nbsp;</TD>
                  <TD width="550">&nbsp;</TD>
                  <TD width="32">&nbsp;</TD>
                </TR>
                <TR>
                  <TD rowspan="2" width="3">&nbsp;</TD>
                  <TD class="bg1" colspan="2" height="38" style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign="top"> 
                    <TABLE bgcolor="#6d4110" border="0" bordercolor="#669900" cellpadding="1" cellspacing="1" width="100%">
                      <TBODY> 
                      <TR bgcolor="#ffffff"> 
                        <TD align="middle" bgcolor="#ffffff"> 
                          <DIV align="center"><font color="#65251C"><%=rs("title")%></font></DIV>
                         
                        </TD>
                      </TR>
                   
                      <TR bgcolor="#c4deff"> 
                        <TD bgcolor="#ffffff" height="73"> 
                          <DIV align="center" style="top : 165px;left : 538px;"><%=HtmlOut(rs("content"))%></DIV>
                        </TD>
                      </TR>
                      <TR bgcolor="#c4deff">
                        <TD bgcolor="#ffffff">&nbsp;</TD>
                      </TR>

                      </TBODY> 
                    </TABLE>

                  </TD>
                </TR>


                  <TD class="bg1" colspan="2" style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign="top" height="10"> 
                    <div align="right"></div>
                  </TD>
                </TR>
              </TBODY>
            </TABLE>
            </TD>
          </TR>
          <TR>
            <TD align="center" bgcolor="#e8e3d0" height="20" valign="bottom"><A class="w" href="javascript:window.history.back()"><B>[返 
              回]</B></A></TD>
          </TR>
        </TBODY>
      </TABLE>
      </TD>
    </TR>
    <TR>
      <TD><IMG src="../pic/a_image51.gif" width="550" height="17" border="0"></TD>
    </TR>
  </TBODY>
</TABLE>
</CENTER>
</BODY>
</HTML>
<%
end if
%>
