<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<html>
<link rel="stylesheet" href="../css.css">
<title>職業介紹所</title>
<body bgcolor="#996699" background="../chat/bk.jpg" leftmargin="0" topmargin="0">
<table border="0" cellspacing="0" cellpadding="0" width="382" align="center">
  <tr> <td height="81" valign="top">
      <form method=POST action='cwork.asp'> 
        <table width="366" BORDER="1" align="center" background="../ff.jpg">
          <tr> <td> <tr> <td align=center>請選擇您的職業： 
<select name=jiu size=1>
                <option value="戰士" selected>戰士</option>
                <option value="銅盔戰士">銅盔戰士</option>
                <option value="銀鎧戰士">銀鎧戰士</option>
                <option value="金甲戰士">金甲戰士</option>
                <option value="神龍戰士">神龍戰士</option>
                <option value="勇士">勇士</option>
                <option value="百戰勇士">百戰勇士</option>
                <option value="虎威勇士">虎威勇士</option>
                <option value="擒龍勇士">擒龍勇士</option>
                <option value="金剛勇士">金剛勇士</option>
                <option value="道士">道士</option>
                <option value="水道士">水道士</option>
                <option value="水法師">水法師</option>
                <option value="水真人">水真人</option>
                <option value="水天師">水天師</option>
              </select> </td></tr> <tr> <td  align=center> 
<input type=submit value=我選擇好了 name="submit"> </td></tr> <tr> 
            <td valign="top" > 
              <p><FONT SIZE="-1" COLOR="#000000">職業是您在江湖要的經濟來源，選擇不同的職業對您的工作會有不同的影響：<br>
                戰士： 
                等級需求0級以上<br>
                <font color="#FF0000">銅盔戰士：等級需求50級以上 殺怪攻擊＊2倍 <br>
                銀鎧戰士：等級需求250級以上 流星5個 殺怪攻擊＊3倍 <BR>
                金甲戰士：等級需求400級以上 流星淚5個 殺怪攻擊＊4倍 <br>
                神龍戰士：等級需求650級以上 流星30個 殺怪攻擊＊5倍 </font> </FONT></p>
              <p><font color="#000000" size="-1">勇士：等級需求0級以上<br>
                <font color="#000099">百戰勇士：等級需求80級以上 殺怪防御＊2倍 <br>
                虎威勇士：等級需求280級以上 流星5個 殺怪防御＊3倍 <br>
                擒龍勇士：等級需求500級以上 流星淚5個 殺怪防御＊4倍 <br>
                金剛勇士：等級需求750級以上 流星45個 殺怪防御＊5倍 </font> </font></p>
              <p><font color="#0033FF" size="-1">道士：等級需求0級以上<br>
                水道士：等級需求90級以上 
                殺怪武功＊1.5倍<br>
                水法師：等級需求320級以上 
                流星5個 
                殺怪武功＊2倍<br>
                水真人：等級需求550級以上 
                流星淚5個 
                殺怪武功＊2.5倍 
                <br>
                水天師：等級需求780級以上 
                流星65個 
                殺怪武功＊3倍<br>
                道士類有補血、武功等功能</font></p>
              <p align="center"><font color="#000000">轉職另收取100萬轉職費</font></p>
            </td></tr> 
</table></form></td></tr> </table>
<div align="center"><font color="#00FF66"><b></b></font> </div>
</body>
</html>
 