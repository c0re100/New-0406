<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 會員等級,會員費,金幣 from 用戶 where id="&info(9),conn
if rs("會員等級")=0 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你不是會員請回吧！');location.href = '../help/info.asp?type=2&name=會員加入辦法';}</script>"
response.end
end if
hy=rs("會員費")
jb=rs("金幣")
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<title>會費轉金幣</title>
<link REL="StyleSheet" HREF="../style.css" TITLE="Contemporary">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" LEFTMARGIN="0" TOPMARGIN="0" >
<div align="center"><center>
    <table border="0" width="513" cellspacing="0" cellpadding="0" ALIGN="CENTER">
      <tr> 
        <td align="center"> 
          <p><br>
            <FONT COLOR="#FF66FF" SIZE="2">您現有會費： </FONT><font color="#000000"><b><font color=#FFFFCC><%=hy%></font></b></font><FONT COLOR="#FF66FF" SIZE="2"> 
            金幣： </FONT><font color="#000000"><b><font color=#FFFFCC><%=jb%> <br>
            轉換概率:1元會費等於100個金幣</font></b></font></p>
          <form name="form1" method="post" action="hxiehyok.asp">
            <input type="text" name="text" value="請輸入你要轉換的會費！" maxlength="10" size="22">
            <input type="submit" name="Submit" value=" 轉 換 ">
          </form>
          <p>&nbsp;</p>
          </td>
      </tr>
      <tr> 
        <td align="center"></td>
      </tr>
    </table>
  </center>
</div>
</body>
</html>
<html>