<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
nowdate=cstr(date())
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �Ȩ�,�s��,������ from �Τ� where id="&info(9),conn
lastdate=rs("������")
if lastdate="" or isnull(lastdate) then lastdate=nowdate
newmoney=clng(rs("�s��")*1.01^datediff("d",lastdate,nowdate))
money=rs("�Ȩ�")
conn.Execute "update �Τ� set �s��="&newmoney&",������='"&nowdate&"' where id="&info(9)
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="../style.css">
<title>�������</title>
<script language=javascript>
function check(){
	if(document.form1.op[0].checked & parseInt(document.form1.mn.value)<parseInt(document.form1.money.value)){
		alert('�A���{�������A�L�k�s�J');
		return false;
	}
	if(document.form1.op[1].checked & parseInt(document.form1.nmn.value)<parseInt(document.form1.money.value)){
		alert('�A���s�ڤ����A�L�k���X');
		return false;
	}
	return true;
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor='#996666' background="../ff.jpg" LEFTMARGIN="0" TOPMARGIN="0" oncontextmenu=self.event.returnValue=false>
<p align=center>&nbsp;</p><form name=form1 method=post action=bankoption.asp onsubmit=return(check())> 
<div align="center"> <center>�z�b�F�C��������s��<FONT COLOR="#FFFFFF"><%=newmoney%></FONT>��!�Q��1�H!<BR><BR> 
      <table width='50%' border=0 cellspacing="2" cellpadding="2" style="border-collapse: collapse">
        <tr bgcolor="#FFCC00"> 
          <td>�{��</td>
          <td> 
            <input type=text name=mn readonly value='<%=money%>' size=9 class=norstyle></td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td>�Ȳ�</td>
          <td> 
            <input type=text name=nmn readonly value='<%=newmoney%>' size=9 class=norstyle></td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td colspan=2 align=center> 
            <input type=radio name=op checked value='�s��'>�s��<input type=radio name=op value='����'>����</td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td colspan=2 > 
            <input name='money' type=text maxlength=9 size=9>���B��J�Ȩ�ƥ�</td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td height="32" colspan=2 align=center> 
            <input type=submit value='����'></td></tr> 
</table></center></div></form></div> <DIV ALIGN="CENTER"><input type="button" value=" �� �� " onClick="javascript:window.close();" name="button"> 
</body>