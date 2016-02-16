<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<HTML><HEAD><TITLE>靈劍江湖-怡紅院</TITLE><META content="text/html; charset=big5" http-equiv=Content-Type><link rel="stylesheet" href="../css.css">
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD><BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0">
<div align="center">::<font size="2" color="#FD794D"><b>怡紅院</b></font>::<br>
  <br>
</div>
<table border="1" cellpadding="0" cellspacing="0" width="554" align="center">
  <tr> 
    <td align="center" colspan="2"> <font size="2" color="#65251C"><b><%=info(5)%>的<%=info(0)%>進入靈劍江湖怡紅院，在這兒，你可以找到美麗的姑娘聊天哦</b></font></td>
  </tr>
  <tr> 
    <td align="center" rowspan="2" width="150"><img src="girl.jpg" width="150" height="174"></td>
    <td width="398"> <font size="3" color="#65251C"><b><font color="#FF0000">韋春花</font>說道：</b></font><font size="2" color="#65251C">大蝦，快到我們怡紅院來，和外我們這漂亮的姑娘聊聊天，你就能增加內力哦，不過我們這兒的姑娘都是賣藝不賣身哦，如果你是姑娘，沒錢的話，你也可以到我們這來工作哦，不過如果你的美貌不夠的話就不來試了，嘻嘻，我們這隻要美女哦。</font></td>
  </tr>
  <tr> 
    <td width="398" valign="middle" align="center" height="107"> 
      <p><font size="2" color="#65251C"><b><a href="xiaojie.asp">人在江湖，沒一個紅粉知己怎麼行呀，我要進怡紅院和踫踫運氣。</a></b></font></p>
      <p><font size="2" color="#65251C"><b><a href="dengji.asp">哎，沒錢了，我要在這上班。<br>
        </a><br>
        <font size="2" color="#65251C"><a href="ss.ASP">我要離開青樓</a></font></b></font></p>
      <p>&nbsp; </p>
    </td>
  </tr>
</table>
</BODY></HTML> 