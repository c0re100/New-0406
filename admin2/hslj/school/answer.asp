<%
bgcolor=Application("Ba_jxqy_backgroundcolor")
Response.Expires=0
if session("Ba_jxqy_username")="" then Response.Redirect "../error.asp?id=016"
Set conn = Server.CreateObject("ADODB.Connection")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & _
Server.MapPath("message.asp")
Set rs = Server.CreateObject("ADODB.Recordset")
sql = "Select * from QUESTION "
rs.Open sql, conn, adOpenStatic
%>
<script language="javascript">
function popwin(id)
{
window.open("edit.asp?id="+id+"",'','menubar=yes,scrollbars=yes resizable=yes width=480,height=300');
}
</script>
<script language="javascript">
function click() {
if (event.button==2) {
alert('PPMM盯著你，像要殺人：不要作弊喲~！！！')
}
}
document.onmousedown=click
</script>
<html>
<head>
<title>翰林試場</title>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<LINK href="../style4.css" rel=stylesheet></head>

<body oncontextmenu=self.event.returnValue=false background="<%=bgimage%>" bgcolor="<%=bgcolor%>"><table border="0" width="100%">
<tr>
<td width="100%">
<p style="line-height: 200%; margin-left: 100; margin-right: 100; margin-top: 10; margin-bottom: 10">這裡是翰林試場，主考官威嚴的看著考生們。隻要花費1000兩銀子就可以參加考試，每答對一道踢就可以加上10兩的銀子、10點的精力和3點的資質。<font color="#FF0000">祝你好運&nbsp;</font></td>
</tr>
</table>
<form method="post" action="answer2.asp" name="form">
<table width="96%" cellspacing="1" cellpadding="0" border="1" bordercolor="#000000" align="center">
<tr>
<td width="82%" height="17"><b>題目</b></td>
<td width="18%" height="17">&nbsp;</td>
</tr>
<%dim i,ran
for i=1 to 10
randomize
ran=(int((6-0+1)*rnd+0))&(int((9-0+1)*rnd+0))&(int((9-0+1)*rnd+0))
ran=clng(ran)
rs.movefirst
rs.move ran-1
If rs.EOF Then Exit For
%>
<tr>
<td width="82%">第<%=I%>題：<%=rs("question")%></td>
<td width="18%">選擇答案
<input type="hidden" name="ann<%=i%>" value="<%=rs("answer")%>">
</td>
</tr>
<tr>
<td width="82%">A：<%=rs("A")%> </td>
<td width="18%">
<input type="radio" name="AN<%=I%>" value="A">
</td>
</tr>
<tr>
<td width="82%">B：<%=rs("b")%></td>
<td width="18%">
<input type="radio" name="AN<%=I%>" value="B">
</td>
</tr>
<tr>
<td width="82%">C：<%=rs("c")%></td>
<td width="18%">
<input type="radio" name="AN<%=I%>" value="C">
</td>
</tr>
<tr>
<td width="82%" height="26">D：<%=rs("d")%></td>
<td width="18%" height="26">
<input type="radio" name="AN<%=I%>" value="D">
</td>
</tr>
<%
Next
%>
</table>
<div align="center">
<input type="submit" name="Submit" value="提交">
</div>
</form>
</body>
</html>
