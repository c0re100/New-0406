<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<!--#include file="jhb.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�жi�J��ѫǦA�i�J�Ѳ��I"
window.close()
</script>
<%Response.End
end if

Jname=info(0)
sql="delete * from ��� where �ɶ�<date()"
conn.execute sql
sql="select * from ��� where �b��='" & Jname & "' "
set rs=conn.execute(sql)

%>
<body bgcolor="#990000"><table border=1 bgcolor="#bee0fc" align=center width=550 cellpadding="10" cellspacing="13"> 
<option><p align=center><font size=5 color=#00ff00></font></p>
	<tr><td bgcolor=#cce8d6>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" >		</table>
		
		<P align=center><front size=3pt>�g���H���Ѭ��z�N�z������p�U�A����1%���Ī��N�O�z�b�ѥ������q</front></p>
		<table border=0 cellpadding=1 cellspacing=1 bgcolor="#51a8ff" width="100%" 
     >
	
		
		
        <P align=center></P>
        <TBODY>
			<tr bgcolor=#c4deff>
				<td>�Ѳ��W��</td><td>�R��ާ@</td><td>����ƶq</td><td>�ɶ�</td><td>���q</td>
			</tr>
<%DO UNTIL RS.EOF%>	
			<tr bgcolor=#f7f7f7>
				<td><%=rs("���~")%>
				<td><%=rs("�ާ@")%></td>
				<td><%=rs("����q")%></td>
				<td><%=rs("�ɶ�")%></td>
				<td><%=formatnumber(rs("����"),0)%></td>
				
		    </tr>
<%
rs.movenext
loop
set rs=nothing
conn.close
set conn=nothing
%></p>

		</table>