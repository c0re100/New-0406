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
rs.open "select 銀兩,存款,結算日期 from 用戶 where id="&info(9),conn
lastdate=rs("結算日期")
if lastdate="" or isnull(lastdate) then lastdate=nowdate
newmoney=clng(rs("存款")*1.01^datediff("d",lastdate,nowdate))
money=rs("銀兩")
conn.Execute "update 用戶 set 存款="&newmoney&",結算日期='"&nowdate&"' where id="&info(9)
rs.Close
set rs=nothing
conn.Close
set conn=nothing
%>
<head>
<link rel="stylesheet" href="../style.css">
<title>江湖錢莊</title>
<script language=javascript>
function check(){
	if(document.form1.op[0].checked & parseInt(document.form1.mn.value)<parseInt(document.form1.money.value)){
		alert('你的現金不足，無法存入');
		return false;
	}
	if(document.form1.op[1].checked & parseInt(document.form1.nmn.value)<parseInt(document.form1.money.value)){
		alert('你的存款不足，無法取出');
		return false;
	}
	return true;
}
</script>
<meta http-equiv="Content-Type" content="text/html; charset=big5"></head>
<body bgcolor='#996666' background="../ff.jpg" LEFTMARGIN="0" TOPMARGIN="0" oncontextmenu=self.event.returnValue=false>
<p align=center>&nbsp;</p><form name=form1 method=post action=bankoption.asp onsubmit=return(check())> 
<div align="center"> <center>您在靈劍江湖錢莊存有<FONT COLOR="#FFFFFF"><%=newmoney%></FONT>兩!利息1％!<BR><BR> 
      <table width='50%' border=0 cellspacing="2" cellpadding="2" style="border-collapse: collapse">
        <tr bgcolor="#FFCC00"> 
          <td>現銀</td>
          <td> 
            <input type=text name=mn readonly value='<%=money%>' size=9 class=norstyle></td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td>銀票</td>
          <td> 
            <input type=text name=nmn readonly value='<%=newmoney%>' size=9 class=norstyle></td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td colspan=2 align=center> 
            <input type=radio name=op checked value='存款'>存款<input type=radio name=op value='取款'>取款</td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td colspan=2 > 
            <input name='money' type=text maxlength=9 size=9>此處輸入銀兩數目</td></tr> 
        <tr bgcolor="#FFCC00"> 
          <td height="32" colspan=2 align=center> 
            <input type=submit value='執行'></td></tr> 
</table></center></div></form></div> <DIV ALIGN="CENTER"><input type="button" value=" 關 閉 " onClick="javascript:window.close();" name="button"> 
</body>