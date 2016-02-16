<%
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>

<html>
<link rel="stylesheet" href="../../css.css">
<body leftmargin="0" topmargin="0" bgcolor="#660000" text="#FFFFFF">
<table border="0" cellspacing="0" cellpadding="0" width="597">
  <tr> 
    <td bgcolor="#660000" align="right" colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td bgcolor="#660000" align="right" width="330" height="328" valign="top"> 
      <p align=center style='color:yellow'><font color="#FF9900">歡迎<%=info(0)%>光臨醉仙樓</font> 
        <br>
      <table border="0" align="center" cellpadding="0" cellspacing="0" width="300" height="200">
        <tr> 
          <td colspan="3"><font color="#FF9900">這裡是京城的名牌酒樓,隻見裡面喧嘩聲不斷好不熱鬧.在酒樓門口有一個牌子上面寫著:</font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> 酒 名 </font></td>
          <td><font color="#FF9900"> 體 力 </font></td>
          <td><font color="#FF9900"> 銀 兩 </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> 老百干 </font></td>
          <td><font color="#FF9900"> 50 </font></td>
          <td><font color="#FF9900"> 250 </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> 五糧液 </font></td>
          <td><font color="#FF9900"> 60 </font></td>
          <td><font color="#FF9900"> 300 </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> 杜康 </font></td>
          <td><font color="#FF9900"> 70 </font></td>
          <td><font color="#FF9900"> 350 </font></td>
        </tr>
        <tr align="center"> 
          <td><font color="#FF9900"> 茅臺 </font></td>
          <td><font color="#FF9900"> 80 </font></td>
          <td><font color="#FF9900"> 400 </font></td>
        </tr>
      </table>
    </td>
    <td align="left" height="328" width="59">&nbsp;</td>
    <td align="center" width=300 height="328" valign="top"> <br>
      <form method=POST action='pub1.asp'>
        <table width="300">
          <tr> 
            <td> 
          <tr> 
            <td align=center>你向小二買下了： 
              <select name=jiu size=1>
                <option value="lao">老百干 
                <option value="wu">五糧液 
                <option value="du">杜康 
                <option value="mao">茅臺 
              </select>
              就舉起酒杯</td>
          </tr>
          <tr> 
            <td  align=center> 
              <input type=submit value=一飲而盡 name="submit">
            </td>
          </tr>
          <tr> 
            <td > 1、在酒樓喝酒可以增加體力； <br>
              2、酒雖然是個好東西,但是不能貪多； 酒量不夠請少飲,否則,嘻嘻......<br>
              3、酒量與你的武功和經驗有關！ 酒量=(武功+經驗)/100 要想不醉倒須 酒量>1， 自己掂量掂量吧！</td>
          </tr>
        </table>
      </form>
    </td>
  </tr>
</table>
</body>
</html>


