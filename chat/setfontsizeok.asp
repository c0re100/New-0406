<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(0)="" or Session("ljjh_inthechat")<>"1" or Instr(useronlinename," "&info(0)&" ")=0 then Response.Redirect "chaterr.asp?id=001"
if chatbgcolor="" then chatbgcolor="008888"%><html>
<head>
<title>�]�m����</title>
<meta http-equiv='content-type' content='text/html; charset=big5'>
<link rel="stylesheet" href="readonly/style.css">
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" leftmargin="0" bgproperties="fixed" topmargin="0">
<div align=center>
<tr>
<td>
      <table border="1" bgcolor="E0E0E0" cellspacing="0" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" cellpadding="4" width="154" height="267">
        <form>
<tr>
            <td style="font-size:10.5pt" bgcolor="#EEFFEE"> 
              <div align=center><font color="#FF0000" style="font-size:12pt">�]�m����</font></div>
<p> �w�g�N�r���]�m�� <font color="#FF0000"><script>document.write(parent.f2.document.af.fs.value);</script>�S</font>�A��Z�]�m�� <font color="#FF0000"><script>document.write(parent.f2.document.af.lh.value);</script>%</font> �A�����I�����M�̡��~��ϸӦr���ͮġC</p>
<div align=center>
<input type="button" value="��^" onclick=javascript:history.go(-1)>
              </div>
</td>
</tr>
</form>
</table>
</td>
</tr>
</div>
</body>
</html> 
