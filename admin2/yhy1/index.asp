<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")%>
<HTML><HEAD><TITLE>靈劍江湖-迎花宮</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="../pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 4.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="bg.jpg">
<div align="center">
  <table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="32%" height="31">
    <tr align="center"> 
      <td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">迎花宮</font></b></td>
    </tr>
  </table>
  <br>
</div>
<table border=0 cellpadding=0 cellspacing=0 width="530" align="center">
  <tbody> 
  <tr> 
    <td class=bg1 colspan=2 rowspan=2 
          style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
      <table border="0" cellpadding="0" cellspacing="0" width="554" height="231">
        <tr> 
          <td align="center" colspan="2"> <font size="2" color="#65251C"><b> <font color="#FFFFFF"><%=info(0)%><font color="#00FF00">進入迎花宮，在這兒，你可以找到酷帥的哥哥聊天哦</font></font></b></font> 
          </td>
        </tr>
        <tr> 
          <td align="center" rowspan="2" width="200"><img src="girl.jpg" width="200" height="151"> 
          </td>
          <td width="343"> 
            <p>&nbsp;</p>
            <p><font size="3" color="#65251C"><b><font color="#FFFFFF">鬼大頭說道：</font></b></font><font size="2" color="#00FFFF">大蝦，快到我們迎花宮來，和我們這超酷的帥男談談心，你就能增加內力哦，不過我們這兒的帥男都是賣藝不賣身哦，如果你是帥哥，沒錢的話，你也可以到我們這來工作哦，不過如果你的美貌不夠的話就不來試了，哈哈，我們這隻要帥哥哦。</font> 
            </p>
          </td><br>
        </tr>
        <tr>
          <td width="343" valign="middle" align="center" height="115"> 
            <p>&nbsp;</p>
            <p><font size="2" color="#65251C"><b><a href="xiaojie.asp"><font color="#0000FF">人在江湖，隨波逐流，衝進迎花宮</font></a></b></font> 
            </p>
            <p><font size="2" color="#65251C"><b><a href="dengji.asp"><font color="#FF00FF">哎，沒錢了，先在這上班算了</font></a><br>
              </b></font></p>
            <p><b><font size="2"><a href="delgirl.asp"><font color="#FFFF00">不干了，這是人過的日子嗎？我要贖身</font></a></font></b></p>
            </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td colspan=2>&nbsp;</td>
  </tr>
  </tbody> 
</table>
</BODY></HTML>       
