<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(5)<>0 then
Response.Write "<script Language=Javascript>alert('�n��սХ��h�j�U�I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if

%>
<html>
<head>
<title>�k�O�䧽</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link rel="stylesheet" href="../../CSS.CSS" type="text/css">
</head>
<body oncontextmenu=self.event.returnValue=false bgproperties="fixed" bgcolor="#660000" topmargin="0" leftmargin="0">
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" height="441" >
  <tr  onMouseOut="this.bgColor='#FFFFFF';"onMouseOver="this.bgColor='#DFEFFF';"> 
    <td bgcolor="#EEFFEE" height="31"> 
      <div align="center"><font color="#FF6633"> �k�O�䧽�U�` </font></div>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FF6633" height="77"> 
      <div align="left"> 
        <p> <font color="#FFFFFF">�k�O�䧽:�̤�100�I�A�̦h5000�I<br>
          �w�I�䧽:�̤�10�I�A�̦h500�I </font> 
      </div>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FF6633" height="57"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">�U�`��j�G</font> <br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> 
          <input type=submit value="�j" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
#000000 solid">
          </font> </div>
      </form>
    </td>
  </tr>
  <tr> 
    <td bgcolor="#FF6633" height="53"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">�U�`��p�G</font> <br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> 
          <input type=submit value="�p" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
#000000 solid">
          </font> </div>
      </form>
  <tr> 
    <td height="57" bgcolor="#FF6633"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">�U�`���G</font><br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> </font><font color=#EEFFEE size=2> 
          <input type=submit value="��" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
#000000 solid">
          </font><br>
        </div>
      </form>
    </td>
  </tr>
  <tr> 
    <td height="57" bgcolor="#FF6633"> 
      <form method=POST action=duju003ok.asp>
        <div align="center"><font color="#FFFFFF">�U�`�����G</font><br>
          <input type="text" name="num" size="10" maxlength="8" style="border: 1px solid;background-color:#FFFFDF; font-size: 9pt; border-color:
#000000 solid">
          <font color=#EEFFEE size=2> 
          <input type=submit value="��" name=submit style="border: 2px solid;background-color:#ffffff; font-size: 9pt; border-color:
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