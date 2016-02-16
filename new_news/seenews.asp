<%
id  = Request("id")
Set conn = Server.CreateObject ("ADODB.Connection")
DBPath = Server.MapPath("new.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath
SET rs   = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From news Where (id=" & id & ")",conn,2,3
RS("count")= RS("count") + 1
RS.Update
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 5.0">
<TITLE>江湖公告</TITLE>
<style type="text/css">
<!--
body, td, form, option, select, textarea, p, ol{font-family:新細明體,Arial;font-size:9pt;line-height:1.6};
A:link {color:#6633ff; text-decoration: none}
A:visited {color: #6633ff; text-decoration: none}
A:hover {color: #ff6600; text-decoration: underline}
-->
</style>
</head>


<HTML>
<HEAD>
<TITLE>江湖公告</TITLE>
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
      <TD><IMG src="img/jh_pic.gif" width="550" height="25" border="0"></TD>
    </TR>
    <TR>
      <TD background="img/a_image41.gif" align="center">
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
                  <TD colspan="2" nowrap>&nbsp; </TD>
                </TR>
                <TR> 
                  <TD rowspan="2" width="2">&nbsp;</TD>
                  <TD class="bg1" height="38" style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign="top"> 
                    <TABLE bgcolor="#6d4110" border="0" bordercolor="#669900" cellpadding="1" cellspacing="1" width="100%">
                      <TBODY> 
                      <TR bgcolor="#ffffff"> 
                        <TD align="middle" bgcolor="#ffffff"> 
                          <DIV align="center"><font color="#65251C"><b><%=rs("title")%></b></font></DIV>
                        </TD>
                      </TR>
                      <TR bgcolor="#c4deff"> 
                        <TD bgcolor="#ffffff"><%=rs("news")%><BR> </TD>
                      </TR>
                      <TR bgcolor="#c4deff"> 
                        <TD bgcolor="#ffffff"> 
                          <div align="right">發 布 人:<font color="#65251C"><b><%=rs("name")%></b></font><br>
                            發佈時間：<font color="#65251C"><b><%=rs("d")%><BR>被閱讀：<%=rs("count")%>次</b></font> 
                          </div>
                        </TD>
                      </TR>
                      </TBODY> 
                    </TABLE>
                  </TD>
                </TR>
                <TD class="bg1" style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign="top" height="10"> 
                  <div align="right"></div>
                </TD>
                </TR>
                </TBODY> 
              </TABLE>
            </TD>
          </TR>
          <TR>
            <TD align="center" bgcolor="#e8e3d0" height="20" valign="bottom"><a class="w" href="javascript:window.close()"><b>[關閉窗口]</b></a></TD>
          </TR>
        </TBODY>
      </TABLE>
      </TD>
    </TR>
    <TR>
      <TD><IMG src="img/a_image51.gif" width="550" height="17" border="0"></TD>
    </TR>
  </TBODY>
</TABLE>
</CENTER>
</BODY>
</HTML>
