<%@ LANGUAGE=VBScript codepage ="950" %>
<%
Set Conn=Server.CreateObject("ADODB.Connection")
Set rs=Server.CreateObject("ADODB.RecordSet")
Connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath("ljjmwb.asa")
Conn.Open connstr
rs.open "SELECT 網吧付款 FROM mess",conn,2,2
%>
<html>
<head>
<title>靈劍網吧加盟</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
</head>

<body bgcolor="#9933CC" text="#000000" background="../2.jpg">
<div align="center">
  <p><b><font color="#0000CC" size="+2">靈劍江湖網吧聯盟<br>
    </font></b></p>
  <p align="left">1.網吧加盟的條件？<br>
    a.具有一定規模的網吧。<br>
    <br>
    b.可以提供優質上網服務.<br>
    <br>
    c.在本地區具備一定影響力的網吧.<br>
    <br>
    d.需向靈劍江湖提供網吧注冊資料！<br>
    <br>
    e.對於IP地址不固定者，需自行修改IP設置。<br>
    <br>
    2.加入收費標準及優惠正策？<br>
    <br>
    a.加盟網吧收費：200元/月(港、奧、臺及其它地區以美元結算200$元/月!)<br>
    (<b><font color="#FF0000">春節期間特價:150/月，港、奧、臺不變！</font></b>) <br>
    <br>
    b.在加盟網吧上網每天可以多領取工資1000萬兩，會費10元、金幣10個,流星一個！<br>
    <br>
    c.在加盟網吧上網，復活後狀態不會丟失！<br>
    <br>
    d.在加盟網吧上網，<b><font color="#0000FF">泡點是平常網吧上網的2倍！ </font></b><br>
    <br>
    e.靈劍江湖將會在頂部滾動廣告支持一周！<br>
    <br>
    f.在加盟網吧上網購物打7折！(會員不變)<br>
    <br>
    <br>
    <br>
    3.加盟辦法：<br>
  </p>
  <table width="75%" border="1" align="left">
    <tr>
      <td bgcolor="#9999CC"><%=rs("網吧付款")%> 
        <%rs.close
set rs=nothing
conn.close
set conn=nothing%>
      </td>
    </tr>
  </table>
  <p align="left"> <br>
  </p>
</div>
</body>
</html>
