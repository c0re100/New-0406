<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
fl=Trim(Request.QueryString("fl"))
if fl<>"右手" and fl<>"左手" and fl<>"盔甲" and fl<>"頭盔" and fl<>"雙腳" and fl<>"裝飾" and fl<>"藥品" and fl<>"毒藥" and fl<>"卡片" then
	Response.Write "<script Language=Javascript>alert('分類查隻能為：右手、左手、盔甲、頭盔、雙腳、裝飾、藥品、毒藥、卡片，請重新選擇！');window.close();</script>"
	Response.End
end if
if info(2)<7 then 
	Response.Write "<script Language=Javascript>alert('你想作什麼滾！');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT * FROM 物品 WHERE 類型='" & fl & "' and 數量>0  order by 擁有者",conn
%>
<html>
<head>
<title>物品管理</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="style.css">
</head>
<body bgcolor="#660000" text="#FFFFFF" leftmargin="0" topmargin="0">
<div align="center"><font color=yellow><b><%=fl%>類物品</b></font>一覽(裝備的物品除外)<font face="幼圓"><a href="javascript:this.location.reload()">刷新</a></font></div>
<br>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">擁有者</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">物品名</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">類型 </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">數量 </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">內力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">體力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">攻擊 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">防御 </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">價錢</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">裝備</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">方式</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("擁有者")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("物品名")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("類型")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("數量")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("內力")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("體力")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("攻擊")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("防御")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("銀兩")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("裝備")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del.asp?id=<%=rs("id")%>"><font color="#0000FF">刪除</font></a></font></div>
      </td>
  </tr>
  <%
rs.movenext
loop
%>
</table>
<%
rs.close
rs.open "SELECT * FROM 交易市場 WHERE 類型='" & fl & "' order by 擁有者,方式",conn
%>
<table border="1" align="center" width="454" cellpadding="1" cellspacing="0" height="31">
  <tr align="center"> 
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">擁有者</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">物品名</font></div>
    </td>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">類型 </font></div>
    <td nowrap width="33" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">數量 </font> </div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">內力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">體力 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">攻擊 </font></div>
    <td nowrap width="40" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">防御 </font></div>
    <td colspan="2" nowrap height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">價錢</font></div>
    </td>
    <td nowrap width="45" height="16"> 
      <div align="center"><font size="-1">方式</font></div>
    </td>
    <td nowrap width="50" height="16"> 
      <div align="center"><font color="#FFFFFF" size="-1">操作</font></div>
    </td>
  </tr>
  <%
do while not rs.eof
%>
  <tr> 
    <td width="45" height="3"> 
      <div align="center"><font color="#FFFFFF" size="-1"><%=rs("擁有者")%></font></div>
    </td>
      <td width="45" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("物品名")%> </font> 
        </div>
      </td>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("類型")%></font> 
        </div>
      <td width="33" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("數量")%> </font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("內力")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("體力")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("攻擊")%></font> 
        </div>
      <td width="40" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("防御")%></font> 
        </div>
      <td colspan="2" height="3"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("銀兩")%> </font> 
        </div>
      </td>
      <td height="3" width="45"> 
        <div align="center"><font color="#FFFFFF" size="-1"><%=rs("方式")%></font></div>
      </td>
      <td height="3" width="50"> 
        <div align="center"><font size="-1"><a href="del1.asp?id=<%=rs("id")%>"><font color="#0000FF">刪除</font></a></font></div>
      </td>
  </tr>
  <%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<tr align="center"> 
  <td nowrap width="45" height="16"> 
    <div align="center"></div>
  </td>
</tr>
<table width="64%" border="1" cellpadding="0" cellspacing="0" align="center">
  <tr> 
    <td height="25"> 
      <div align="center">這裡可以查看到對方的物品，刪除就可以將此用戶的物品刪除！</div>
    </td>
  </tr>
</table>
</body>
</html>
 