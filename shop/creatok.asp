<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%><!--#include file="data.asp"--><%
sql="SELECT * FROM 商店 where 店主='"&info(0)&"'"
set rs1=conn1.Execute (sql)
if (not rs1.EOF) or (not rs1.BOF) then
set rs1=nothing
conn1.close
set conn1=nothing%>
<script language="vbscript">
alert("你已經開設了商店了！")
window.close()
</script>
<%Response.End 
end if
rs1.close
set rs1=nothing
rs.open "select 會員等級,等級,銀兩 FROM 用戶 WHERE id="&info(9),conn
if rs("會員等級")=0 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，商店老板目前隻對會員開放！');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
if rs("銀兩")<200000000 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，你需要帶齊2億銀兩作為商店的注冊資金！');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
if rs("等級")<100 then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，開設商店需要你的戰鬥級在100級以上！');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
rs.close
rs.open "select id FROM 物品 WHERE (類型='鮮花' or 類型='右手' or 類型='左手' or 類型='藥品' or 類型='裝飾' or 類型='盔甲') and  數量>0 and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('對不起，開設商店需要你有6種以上物品！');location.href = 'javascript:window.close()';</script>"
	Response.End
end if
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
%>
<head>
<title>開設商店</title>
</head>

<body bgcolor="#99CCFF" background="../linjianww/0064.jpg">
<p align="center"><font color="#800080" size="6">開設商店</font></p>
<form method="POST" action="creatok1.asp">
  <div align="center"> <center> 
      <table border="1" width="70%" cellspacing="0" cellpadding="0" bordercolorlight="#000000" bordercolordark="#FFFFFF">
        <tr> 
          <td width="39%">商店名:</td>
          <td width="66%"> 
            <input type="text" name="shopname" size="10" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          </td>
        </tr>
        <tr> 
          <td width="39%">說&nbsp; 明:</td>
          <td width="66%"> 
            <input type="text" name="memo" size="33" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          </td>
        </tr>
        <tr> 
          <td width="39%">經營項目:</td>
          <td width="66%"> 
            <select size="1" name="shoptype" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
              <option value="請選擇" selected>請選擇</option>
              <option value="物品">物品</option>
            </select>
          </td>
        </tr>
        <tr> 
          <td width="105%" colspan="2"> 
            <p align="center"> 
              <input type="submit" value="提 交" name="B1">
          </td>
        </tr>
      </table>
    </center></div></form>
