<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "請進入聊天室再進入當鋪操作！！"
window.close()
</script>
<%response.end
end if%>
<html>
<head>
<title>當鋪</title>
<link rel="stylesheet" href="../../css.css">
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor="#660000" background="../../ff.jpg" text="#000000" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center">
  <table width="491" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td align="center"> <table width="491" align="center" cellspacing="0" border="1"
cellpadding="0">
          <tr> 
            <td bgcolor="#FFCC00"> 
              <div align="center"><font color="#FF6600"> 
                <%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
name=info(0)
rs.open "SELECT 物品名,類型,數量,銀兩,id FROM 物品 WHERE 類型<>'卡片' and 裝備=false and 擁有者=" & info(9) & "  order by 銀兩 ",conn
if rs.bof and rs.eof then
%>
                當鋪老板驚叫道： <br>
                天哪</font>––:-O<br>
                <%=name%>你身上一無所有，不會是想當人吧！<br>
                <%
else
%>
                <%=name%> 你的可當物品有： </div></td>
          </tr>
        </table>
        <table border="0" cellspacing="2" cellpadding="2" width="491" align="center">
          <tr align="center" bgcolor="#CC0000"> 
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">物 
              品 名</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">數 
              量</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">原 
              價</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">現 
              價</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">當出數量</font></td>
            <td style="border-bottom: 1px solid rgb(250,50,50)"><font color="#FFFFFF">操 
              作</font></td>
          </tr>
          <%
do while not rs.bof and not rs.eof
%>
          <tr bgcolor="#FFFFFF"  onmouseout="this.bgColor='#FFFFFF';"onmouseover="this.bgColor='#DFEFFF';"> 
            <form method=POST action='dan1.asp?id=<%=rs("id")%>'>
              <td><%=rs("物品名")%></td>
              <td><%=rs("數量")%></td>
              <td><%=rs("銀兩")%>/個</td>
              <td><%=rs("銀兩")/2%>/個</td>
              <td> <select name="sl"  style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:#000000 solid">
                  <option value="1" selected>1</option>
                  <option value="2">2</option>
                  <option value="3">3</option>
                  <option value="4">4</option>
                  <option value="5">5</option>
                  <option value="6">6</option>
                  <option value="7">7</option>
                  <option value="8">8</option>
                  <option value="9">9</option>
                </select> </td>
              <td><div align="center"> 
                  <input type="SUBMIT" name="Submit" value="進行">
                </div></td>
            </form>
          </tr>
          <%
rs.movenext
loop
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
        </table></td>
    </tr>
  </table>
</div>
</body>
</html>