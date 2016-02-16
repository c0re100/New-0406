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
rs.open "Select 銀兩 from 用戶 where id="&info(9),conn
yin=rs("銀兩")
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<HTML>
<HEAD>
<title>賭大小</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">

</HEAD>
<body text="#000000" link="#0000FF" alink="#0000FF" vlink="#0000FF" leftmargin="0" topmargin="0" background="../images/8.jpg">
<div align="left"></div>
<div align=center> 
  <p><font color="#000000" size="+3">賭大小</font><font size="2"><br>
    最少下注銀子是 <b><font color="#CC0000">10 </font> 兩<br>
    </b>最大下注銀子是 <font color="#CC0000"><b>3000</b></font><b> 兩</b> <br>
    你現在一共有銀子 <b><font color="#CC0000"><%=yin%></font> 兩</b> 可以作為賭注</font></p>
  <table width="545" border="0" cellspacing="0" cellpadding="0" align="center">
    <tr> 
      <td> 
        <form method="POST" action="b&amp;spose.asp">
          <table border=1 cellspacing=0 cellpadding=0 align="center" width="350" bgcolor="#CCCCCC" bordercolorlight="#000000" bordercolordark="#FFFFFF">
            <tr align="center" bgcolor="#FFFFFF"> 
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
              <td width="33%"><font size="2" color="#000000"><img src="../jhimg/bbs/run.gif" width="38" height="36"></font></td>
            </tr>
            <tr bgcolor="#336633"> 
              <td width="960" colspan="3" height="29"><font size="2" class="c" color="#000000"><b>&nbsp;&nbsp;<font color="#FFFFFF">請選擇</font></b></font></td>
            </tr>
            <tr bgcolor="#FFFFFF"> 
              <td align=center colspan="3"> 
                <table border="0" cellspacing="0" cellpadding="0" bgcolor="#FFFFFF">
                  <tr align="center"> 
                    <td width="50%"><img src="../jhimg/bbs/big.gif" width="46" height="40"></td>
                    <td width="50%"><img src="../jhimg/bbs/small.gif" width="46" height="40"></td>
                  </tr>
                  <tr align="center"> 
                    <td width="50%"> 
                      <input type="radio" name="select" value="big" checked>
                    </td>
                    <td width="50%"> 
                      <input type="radio" name="select" value="small">
                    </td>
                  </tr>
                </table>
              </td>
            </tr>
            <tr> 
              <td align=center colspan="3"><font size="2" color="#000000">我要下注： 
                <input type="text" name="splosh" size="4" value="0" maxlength="4">
                &nbsp;<b>兩</b></font></td>
            </tr>
            <tr> 
              <td align=center colspan="3"> <font size="2" color="#000000"> 
                <input type="submit" value="下注啦！！！" name="B1" >
                <input type="reset" value="我要考慮一下：）" name="B2" >
                </font></td>
            </tr>
          </table>
        </form>
        <p align="center"><font size="2">警告：每次賭錢，你的魅力值會扣 1 ！</font></p>
      </td>
    </tr>
  </table>
<p align="center">靈劍江湖版權所有   </p>
</div>
</BODY>
</HTML>
