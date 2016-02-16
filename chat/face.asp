<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%response.expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%><html>
<head>
<title>更改頭像</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<style type='text/css'>
body{font-size:9pt;}
td{font-size:9pt;}
input{font-size:9pt;}
a{font-size:9pt; color:black;text-decoration:none;}
a:hover{color:red;text-decoration:none;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgproperties="fixed" leftmargin="0" topmargin="20" text="#990033" bgcolor="#660000"  background=bk.jpg>
<form method="POST" action="changeface.asp"> 
  <div align="center"><font size="+1"><b><font color="#FFFF00">更改頭像</font></b></font><font color="#FFFF00" size="+1"><br>
    </font> 
    <table width="118" border="0">
      <tr>
        <td bgcolor="#EEFFEE"><font color="#FFFF00"><img id=face src="../ico/<%=info(10)%>-2.gif" width='114' height='114'  alt=個人形像代表>
<select name=face onChange="document.images['face'].src='../ico/'+options[selectedIndex].value+'-2.gif';" style="BACKGROUND-COLOR:#EEFFEE; BORDER-BOTTOM: 1px double; BORDER-LEFT: 1px double; BORDER-RIGHT: 1px double; BORDER-TOP: 1px double; COLOR: #FF6633" size=1>
            <%for i=0 to 48%>
            <option value='<%=i%>' selected><%=i%></option>
            <%next%>
          </select>
          </font></td>
      </tr>
    </table>
    <font color="#FFFF00"><br>
    </font><font color="#FFFF00"><br>
  <br>
  <br>
  <br>
  <input type="submit" value="修改" name="B1" tyle="font-family: Tahoma; font-size: 12px">
  </font></div>
</form>
</body>
</html> 