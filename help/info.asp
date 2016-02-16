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

news_type=request("type")
name=request("name")
if (news_type="") then
	news_type=3
       name="站長公告"
end if 
Set rs=Server.CreateObject("ADODB.Recordset")
rs.PageSize =pagesize
sql="select * from hc_news where type='"&news_type&"' order by filltime desc"
rs.Open sql,Conn,1,1

maxpage=rs.PageCount 
result_num=rs.RecordCount

if result_num<=0 then
	if search="" then
		word="目前還沒有記錄!"
	else	
		word="沒有查到符合條件的記錄!"
	end if	
else
	maxpage=rs.PageCount 
	page=request("page")
	
	if Not IsNumeric(page) or page="" then
		page=1
	else
		page=cint(page)
	end if
	
	if page<1 then
		page=1
	elseif  page>maxpage then
		page=maxpage
	end if
	
	rs.AbsolutePage=Page
end if
%>

<HTML>
<HEAD>
<TITLE>港龍江湖</TITLE>
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
<meta http-equiv="Content-Type" content="text/html; charset=big5"></HEAD>
<BODY bgcolor="#e8e3d0" background="../bkimg/b_danka1_2.gif"  leftmargin="0" topmargin="0" rightmargin="0">
<CENTER>
  <TABLE width="97%" border="0" cellpadding="2" cellspacing="0">
    <TBODY>
      <TR> 
        <TD  align="center"> <TABLE width="100%" border="0" cellpadding="0" cellspacing="1">
            <TBODY>
              <TR> 
                <TD rowspan="2" width="3">&nbsp;</TD>
                <TD width="582" height="38" valign="top" class="bg1" style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px"> 
                  <TABLE bgcolor="#6d4110" border="0" bordercolor="#669900" cellpadding="1" cellspacing="1" width="100%">
                    <TBODY>
                      <TR bgcolor="#ffffff"> 
                        <TD align="middle" bgcolor="#ffffff" width="4%"> <DIV align="center"></DIV></TD>
                        <TD align="middle" bgcolor="#ffffff" width="74%"><font size="2" color="#65251C">信息主題</font></TD>
                        <TD align="middle" bgcolor="#ffffff" width="22%"> <DIV align="center"><font size="2" color="#65251C">加入時間</font></DIV></TD>
                      </TR>
                      <%
if (result_num>0) then 
	for i=1 to pagesize

%>
                      <TR bgcolor="#c4deff"> 
                        <TD bgcolor="#ffffff" width="4%"> <DIV align="center"></DIV>
                          <DIV align="center"></DIV>
                          <div align="center"></div>
                          <DIV align="center"><img src="../bkimg/ccc.gif" width="11" height="11"></DIV></TD>
                        <TD bgcolor="#ffffff" width="74%"><font color="#65251C"><a href="dis_tg.asp?id=<%=rs("id")%>"><%=rs("title")%></a></font></TD>
                        <TD bgcolor="#ffffff" width="22%"> <DIV align="center" style="top : 165px;left : 538px;"><%=FormatDT(rs("filltime"),8)%></DIV></TD>
                      </TR>
                      <%
	rs.movenext
	if rs.EOF then Exit For
	next
end if

rs.close 	         
set rs=nothing   
conn.close 
set conn=nothing 
%>
                    </TBODY>
                  </TABLE></TD>
              </TR>
            <TD class="bg1" style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign="top" height="10"> 
              <div align="right">江湖共有<b><%=name%> </b><%=result_num%>篇， 
                <%
prop="s=&type="&good_type

maxpage=int(maxpage)
page=int(page)
call ShowLastNext (maxpage,page,prop)
%>
              </div></TD>
            </TR>
          </TABLE></TD>
      </TR>
      <TR> 
        <TD align="center" height="20" valign="bottom"><A class="w" href="javascript:window.close()">※關閉窗口※</A></TD>
      </TR>
  </TABLE>
</CENTER>
</BODY>
</HTML>
