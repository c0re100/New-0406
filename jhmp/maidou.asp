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
rs.open "select 泡豆點數 from 用戶 where id="&info(9),conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=210"
end if
if rs("泡豆點數")<1000 then
	mess="你現在還沒有豆子，只有:<font color=red>"&rs("泡豆點數")&"</font>點"
else
	dousl=int(rs("泡豆點數")/1000)
	conn.execute "update 用戶 set 泡豆點數=泡豆點數-"& dousl*1000 &",會員費=會員費+"& 2*dousl &" where id="&info(9)
	mess="賣豆子："&dousl&"個，會員費增加："&dousl*2&"元，還剩:"&rs("泡豆點數")-dousl*1000&"點！"
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<html>
<link rel="stylesheet" href="../css.css">
<title>賣豆子</title>
<body leftmargin="0" topmargin="0" bgcolor="#003366" text="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="0" width="259" align="center">
  <tr>
<td height="81" valign="top">
      <div align="center"><font color="#000000"><b><font color=#FFCC55><%=info(0)%></font><font color="#CCFF33">歡迎光臨自由市場賣豆</font></b></font></div>
      <table width="280" align="center">
        <tr> 
            <td> 
          <tr> 
            
          <td valign="top" height="11" > 
            <div align="center">豆子簡介</div>
            </td>
          </tr>
          <tr> 
            
          <td valign="top" height="150" > 
            <p>豆子是由你泡點時攢下來了,與泡點另算。每泡夠1000點，計算機系統自動給你一個豆，當你選擇暴豆後，殺傷力為平時3倍，可以持續1小時時間。如果你不想使用，可以在這裡賣了。一個豆子=2元會員費用，你可以用賣豆子的會員費購買只有會員纔有的卡片！</p>
<%=mess%><br><br><br>
            </td>
          </tr>
        </table>

</td>
</tr>
</table>
<div align="center"> </div>
</body>
</html>



