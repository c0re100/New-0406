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
<TITLE>港龍江湖信息中心</TITLE>
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
      <TD><IMG src="../pic/jh_pic.gif" width="550" height="25" border="0"></TD>
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
                  <TD colspan="3" nowrap>&nbsp; </TD>
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
                        <TD bgcolor="#ffffff"> 　　<%=HtmlOut(rs("content"))%> </TD>
                      </TR>
                      <TR bgcolor="#c4deff">
                        <TD bgcolor="#ffffff">
                          <div align="right">發　布　人：<font color="#65251C"><b><%=rs("author")%></b></font><br>
                            發佈時間：<font color="#65251C"><b><%=rs("filltime")%></b></font> 
                          </div>
                        </TD>
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
