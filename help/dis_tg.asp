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
TD{font-size : 12px;font-weight : lighter;line-height : 16px;color : #643706;margin-top : 1px;margin-left : 1px;margin-right : 1px;margin-bottom : 1px;}
A{color : #553006;text-decoration : none;}
</STYLE>
<link rel="stylesheet" href="../images/style.css" type="text/css">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></HEAD>
<BODY bgcolor="#e8e3d0" background="../bkimg/b_danka1_3.gif" leftmargin="0" topmargin="0" rightmargin="0">
<CENTER>
  <br>
  <TABLE bgcolor="#6d4110" border="0" bordercolor="#669900" cellpadding="1" cellspacing="1" width="95%">
    <TBODY>
      <TR bgcolor="#ffffff"> 
        <TD align="middle" bgcolor="#ffffff"> <DIV align="center"><font color="#65251C"><%=rs("title")%> 
            </font></DIV></TD>
      </TR>
      <TR bgcolor="#ffffff"> 
        <TD align="middle" bgcolor="#ffffff"><div align="center"><font color="#65251C"><a class="w" href="javascript:window.close()">※關閉窗口※</a> 
            &nbsp;※<a href="javascript:history.back()">回到上頁</a>※</font></div></TD>
      </TR>
      <TR bgcolor="#c4deff"> 
        <TD bgcolor="#ffffff"><font color="#65251C"><br>
          </font><%=rs("content")%> <font color="#65251C">&nbsp; </font></TD>
      </TR>
      <TR bgcolor="#c4deff"> 
        <TD bgcolor="#ffffff"> <div align="right"> 
            <p>發 布 人:<font color="#65251C"><%=rs("author")%></font><br>
              發佈時間：<font color="#65251C"><%=rs("filltime")%></font> </p>
            <p align="center"><a class="w" href="javascript:window.close()">※關閉窗口※</a> 
              &nbsp;※<a href="javascript:history.back()">回到上頁</a>※ </p>
          </div></TD>
      </TR>
    </TBODY>
  </TABLE>
</CENTER></BODY></HTML>
<%
end if
%>
