<!--#include file="config.asp"-->
<html>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<head>
<style type="text/css">
<!--
a {  text-decoration: none}  
a:hover {  text-decoration: underline} 
table {  font-size: 9pt}
body,table,p,td,input {  font-size: 9pt} 
-->
</style>
<title><%=title%></title>
</head>

<body bgcolor="<%=bgcolor%>" text="<%=textcolor%>" link="<%=linkcolor%>">

<form method="POST" action="cklogin.asp">
      <input type="hidden" name="work" value="manageposted">
       <div align="center"><center><table border="1" cellspacing="1"
    width="400" bordercolor="#C0C0C0">
        <tr>
            
          <td align="center" bgcolor="#0000FF"><font color="#FFFF00">【<%=title%>】板主登錄</font></td>
        </tr>
        <tr>
            <td align="center" style=padding:4>
            板主名字 : <input name="username" maxLength="20" size="15" tabIndex="1"><br>
            板主密碼 : <input type="password" name="password" size="15" tabindex="2"></td>
			</tr>
        <tr>
            <td align="center" height="25"><input
            type="submit" style="border:1 solid #C0C0C0;background-color:white" value="進   入"></td>
        </tr>
      </table>
      </center></div>
    </form>
    </td>
  </tr>
</table>
<!--#include file="copyright.asp"-->
</body>
</html>
