<!--#include file="data2.asp"-->
<html>

<head><meta http-equiv="cnntent-Type" cnntent="text/html; charset=big5"> 
<title>給訪客留言</title>
<link rel="stylesheet" type="text/css" href="../style.css">
<style type="text/css">td           { font-family: 新細明體; font-size: 9pt }
body         { font-family: 新細明體; font-size: 9pt }
select       { font-family: 新細明體; font-size: 9pt }
a            { color: #FFC106; font-family: 新細明體; font-size: 9pt; text-decoration: none }
a:hover      { color: #cc0033; font-family: 新細明體; font-size: 9pt; text-decoration: 
               underline }
</style>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body bgcolor="#000000" text="#FFFFFF">

<table width="744" border="0" cellspacing="0" cellpadding="0" align="center"
height="89">
  <tr>
    <td width="744" background="../jh/tiao20.gif" height="83">
      <table border="0" height="24" width="90%" cellspacing="0" cellpadding="0"
      align="center">
        <tbody>
          <tr>
            <td height="38" width="100%">
              <table width="100%" border="0" cellspacing="0" cellpadding="0"
              bordercolorlight="#000000" bordercolordark="#FFFFFF"
              align="center">
                <tr>
                  
                <td width="91" height="26"><font size="2">&nbsp; <font size="2" color="#FFCC00"><span class="zilong"> 
                  <%
y=Month(date())
r=Day(date())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
Response.Write  y & "月" & r & "日" %> </span></font></font></td>
                  <td width="475" height="26">
                    <div align="center">
                      <font size="2" color="#E18C59"><b> 給訪客留言</b></font>
                    </div>
                  </td>
                  
                <td width="104">&nbsp; </td>
                </tr>
              </table>
            </td>
          </tr>
        </tbody>
      </table>
    </td>
  </tr>
</table>
<table width="738" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="17" background="../jh/tiao10.gif">&nbsp;</td>
    <td valign="top" width="639">
      <div align="center" style="width: 689; height: 41">
        <div align="center"> </div>
        <div align="center">
          <center>
          <table border="1" width="695" cellspacing="1" cellpadding="0"
          bordercolor="#E18C59">
            <tr>
                <td class="p1" width="336">□給訪客留言</td>
              <td class="p1" width="440">□-當前時間：<%=date%><%=time%></td>
            </tr>
          </table>
          </center>
        </div>
        <div align="center">
          <center>
          <form method="POST" action="setguest_post.asp">
            <table border="1" width="690" cellspacing="1" cellpadding="0"
            bordercolor="#E18C59">
              <tr>
                <td class="p3" width="10"> </td>
                <td class="p3" width="326">給來訪者的話<br>
                  <textarea name="guestmemo" rows="7" cols="43"></textarea>
                <td class="p3" width="440" valign="bottom"><input type="submit"
                  value="提交" name="B1"><input type="reset"
                  value="全部重寫" name="B2"></td>
              </tr>
            </table>
          </form>
          </center>
        </div>
        <div align="center">
          <a href="javascript:history.back(1)"><font color="#FFCC00">返回</font></a><center>
          <table border="0" width="690" cellspacing="1" cellpadding="0">
            <tr>
              <td class="p1" width="690">&nbsp;</td>
            </tr>
          </table>
          </center>
        </div>
      </div>
    </td>
    <td width="25" background="../jh/tiao10.gif"> </td>
  </tr>
</table>
<table width="731" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td>
      <div align="right">
        <img src="../jh/tiao19.gif" width="728" height="31">
      </div>
    </td>
  </tr>
</table>
</body>

</html>
