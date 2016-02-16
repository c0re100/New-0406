<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(5)<>0 then
Response.Write "<script Language=Javascript>alert('要賭博請先去大廳！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

%>
<html>
<head>
<title>法力賭局</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../../CSS.CSS" type="text/css">
</head>
<body oncontextmenu=self.event.returnValue=false bgproperties="fixed" bgcolor="#660000" topmargin="0" leftmargin="0">
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" height="441" >
  <tr  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td bgcolor="#EEFFEE" height="31"> 
      <div align="center"><font color="#FF6633"> 法力賭局下注 </font></div>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FF6633" height="77"> 
      <div align="left"> 
        <p> <font color="#FFFFFF">法力賭局:最少100點，最多5000點<br>
          泡點賭局:最少10點，最多500點 </font> 
      </div>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FF6633" height="57"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">下注押大：</font> <br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> 
          <input type=submit value="大" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
#000000 solid">
          </font> </div>
      </form>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FF6633" height="53"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">下注押小：</font> <br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> 
          <input type=submit value="小" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
#000000 solid">
          </font> </div>
      </form>
  <tr> 
    <td height="57" bgcolor="#FF6633"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">下注押單：</font><br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> </font><font color=#EEFFEE size=2> 
          <input type=submit value="單" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
#000000 solid">
          </font><br>
        </div>
      </form>
    </td>
  </tr>
  <tr> 
    <td height="57" bgcolor="#FF6633"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">下注押雙：</font><br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> 
          <input type=submit value="雙" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
#000000 solid">
          </font><br>
        </div>
      </form>
    </td>
  </tr>
</table>
  <p>&nbsp;</p>
  <p class=p1 align="center">&nbsp;</p>
  
</body>
</html>
<html><script language="JavaScript">                                                                  </script></html>