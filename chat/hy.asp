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
sql="select * from 文本 "
Set Rs=conn.Execute(sql)
%><html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<style type="text/css">
A {COLOR: #000000; TEXT-DECORATION: none; TEXT-TRANSFORM: none}
A:hover {COLOR:#C46200; TEXT-DECORATION: underline}
BODY {FONT-FAMILY: "新細明體"; FONT-SIZE: 9pt}
TD {FONT-FAMILY: "新細明體"; FONT-SIZE: 9pt}
</style>
<title>靈劍江湖總站會員簡介</title>
</head>

<body topmargin="0" leftmargin="0" bgcolor="#3366CC" text="#FFFFFF">
<br>
<table border="0" width="375" cellspacing="1" cellpadding="1" align="center" bgcolor="#0099FF">
  <tr> 
    <td width="100%"> 
      <p align="center"><span style="FONT-SIZE: 14px"><b><font color="#FFFF00">靈劍江湖會員簡介（2002.8.10)</font></b></span> 
    </td>
  </tr>
</table>
<br>
<table border="0" width="619" cellspacing="0" cellpadding="0" height="427" align="center" bgcolor="#3366CC">
  <tr> 
    <td width="1" height="897">&nbsp;</td>
    <td width="617" height="897"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>一、為什麼要開始會員功能？</b></font></td>
        </tr>
      </table>
      <p>現在好的空間很難找，要給大家一個好的娛樂環境就必須得有一個好的服務器，現在服務器租用價格很昂貴，經過了多次努力仍然沒有解決經費上的問題，是沒有辦法而為之！會員的操作是在不影響所有玩家的正常使用前題下進行的！<font color="#FFFFFF"><font color="#000000"><br>
        </font></font></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00"><font size="3">二、 會員待遇？</font></font></td>
        </tr>
      </table>
      <p><font color="#FFFFFF"><font color="#000000"></font></font>會員分為4種，具體如下：<br>
        <br>
      </p>
      <table width="74%" border="1" cellspacing="0" cellpadding="2" align="center" height="74" bgcolor="#9966CC" bordercolor="#9966CC">
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FF99FF"><b><font color="#FFFFFF" size="2">1級會員</font><font color="#FFFFFF">：</font></b></font><font color="#FFFF00"><b><%=rs("會員月付1")%>/月</b></font></font></font></font></font>，送銀兩1000萬，加戰鬥等級<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFF00"><%=rs("會員等級1")%>級</font></b></font></font></font></font>，<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00">泡點數：<%=rs("會員泡點1")%>點</font></font></font></font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"></font></font></font></font></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b>2級會員：</b></font><font color="#FFFF00"><b><%=rs("會員月付2")%>元/月</b></font></font></font></font></font>，送銀兩2000萬，加戰鬥等級<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFF00"><%=rs("會員等級2")%>級</font></b><font color="#FFFFFF">，<font color="#000000"><font color="#FFFFFF"><font color="#FFFF00">泡點數：<%=rs("會員泡點2")%>點</font></font></font></font></font></font></font></font></font></font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"> 
              </font></font></font></font></font></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b>3級會員：</b></font><font color="#FFFF00"><b><%=rs("會員月付3")%>元/月</b></font></font></font></font></font>，送銀兩5000萬，加戰鬥等級<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("會員等級3")%>級</b></font><font color="#FFFFFF">，<font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFF00">泡點數：<%=rs("會員泡點3")%>點</font></font></font></font></font></font></font></font><font color="#000000"></font></font></font></font></div>
          </td>
        </tr>
        <tr> 
          <td> 
            <div align="center"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b>4級會員</b><b>：</b></font><font color="#FFFF00"><b><%=rs("會員月付4")%>元/月</b></font></font></font></font></font>，送銀兩8000萬，<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF">加戰鬥等級<font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("會員等級4")%>級</b></font></font></font></font></font></font></font></font>，<font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFF00">泡點數：<%=rs("會員泡點4")%>點</font></font></font></font></font></font></div>
          </td>
        </tr>
      </table>
      <p> <font size="3"> <font size="2" color="#FFFF00">會員購買藥品、裝備時打5折、離開門派不會清銀兩、內力等，名字以特別方式顯示</font></font></p>
      <p><font size="3"><font size="2" color="#FFFF00">1級會員：會費=500元 金幣=金幣+500個<br>
        2級會員：會費=1000元 </font><font size="3"><font size="2" color="#FFFF00">金幣=金幣+1000個</font></font><font size="2" color="#FFFF00"><br>
        3級會員：會費=1500元 </font><font size="3"><font size="2" color="#FFFF00">金幣=金幣+1500個</font></font><font size="2" color="#FFFF00"><br>
        4級會員：會費=2000元 </font><font size="3"><font size="2" color="#FFFF00">金幣=金幣+2000個</font></font></font></p>
      <p><font size="3"><font size="2" color="#FFFF00"> 會費、金幣可以用來買<a href="../card/card.asp" target="_blank"><font color="#00FFFF">會員卡片</font></a>,被偷錢小於10%(江湖為30%)。</font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#99FF99" size="3"><br>
        <br>
        </font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF">靈劍江湖會員分月付、半年付、年付！<b><font color="#FFFF00"><br>
        <font size="+1"> </font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("半年付")%></b></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font><b><font color="#FFFF00"></font></b></font></font></font></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>三、單項購買<font size="3"> 
            </font></b></font></td>
        </tr>
      </table>
      <p>戰鬥級：100元30級<br>
        銀 兩：10元5000萬<br>
        會 費：10元20兩</p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>四、終身會員<font size="3"> 
            </font></b></font></td>
        </tr>
      </table>
      <p><font size="3"><font size="2" color="#00CCFF"><b><font color="#00FFFF">終身會員(3級):</font><font color="#FFFF00">150元，送戰鬥等級40級，銀兩5000萬，泡點數：12點 
        （注：終身會員隻屬於其本人使用有效，不得外借或送人，否則暫停其帳號！） <br>
        隻交一次，以後不用再交<br>
        終身會員介紹一個終身會員獎勵戰鬥級10級、銀兩1000萬、會費10元</font></b></font></font><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><br>
        <br>
        </font></font>無論是交費會員還是其它網管、掌門，如有違反江湖規定，將永久不得錄用，因被開除的會員我們不退回購買時交納的費用，請大家注意！</font><br>
        <font color="#FF0000"><b><font color="#FFFF00"><br>
        </font></b></font></p>
      <table width="100%" border="0" cellspacing="0" cellpadding="0" bgcolor="#0099FF">
        <tr> 
          <td bgcolor="#CC6600"><font color="#FFFF00" size="3"><b>五、付款方法<font size="3"> 
            </font></b></font></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      <p><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("會員龍卡")%></b></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></p>
      <p><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><b><font color="#FFFF00"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><b><font color="#FFFFFF"><font color="#000000"><font color="#FFFFFF"><font color="#000000"><font color="#FFFF00"><b><%=rs("會員地址")%></b></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></font></font></font></font></b></font></font></font></font></font></font></font></font></font></b></font></font></font></font></b></font></font></font></font></font></font></p>
    </td>
    <td width="10" height="897">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1" height="22">&nbsp;</td>
    <td width="617" height="22"> 
      <p align="center"><a href="javascript:window.close()"><font color="#FFFFFF">關閉窗口</font></a> 
    </td>
    <td width="10" height="22">&nbsp;</td>
  </tr>
  <tr> 
    <td width="1" height="15"></td>
    <td width="617" height="15"></td>
    <td width="10" height="15"></td>
  </tr>
  <tr> 
    <td width="1" bgcolor="#000099" height="2"></td>
    <td width="617" bgcolor="#000099" height="2"></td>
    <td width="10" bgcolor="#000099" height="2"></td>
  </tr>
</table>  
  
</body>  
  
</html>  
<%
	set rs=nothing	
	conn.close
	set conn=nothing %>