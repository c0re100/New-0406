<%@ LANGUAGE = VBScript codepage ="950" %>
<%
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if session("TWT_ARR_ArgALL")="" then response.end
TWT_ArrArg=split(session("TWT_ARR_ArgALL"),"=")
nickname=TWT_ArrArg(0)
grade=TWT_ArrArg(2)
myid=TWT_ArrArg(1)
if grade>10 then
sql="delete * FROM �Τ� where id=" & myid
set Rs=conn.Execute(sql) 
conn.close 
session.Abandon 
Response.write "�n�n��,�h���a!�ڭ̤��w��p�O�Ӷ«ȡC" 
response.end 
end if
%>
<HTML><HEAD><TITLE>�ѥ~�Ѧ���--�x���Ū�</TITLE>
<META content="text/html; charset=big5" http-equiv=Content-Type><LINK 
href="pic/css.css" rel=stylesheet>
<META content="Microsoft FrontPage 5.0" name=GENERATOR></HEAD>
<BODY bgColor=#fffddf leftMargin=0 topMargin=0 marginheight="0" marginwidth="0" background="pic/bg.jpg">
<table border=1 cellspacing=0 cellpadding=2 align="center" bordercolordark="#FFFFFF" width="32%" height="31">
<tr align="center"> 
<td bgcolor="#669900" width="100%" height="25"><b><font size="4" color="#FFFFFF">�ѥ~�ѸŪ�</font></b></td>
</tr>
</table>
<br>
<table border=0 cellpadding=0 cellspacing=0 width="601" align="center">
<tbody> 
<tr> 
<td width="11"><img src="pic/empty.gif" width="1" height="1"></td>
<td noWrap colspan="3">&nbsp; </td>
</tr>
<tr> 
<td width="11">�@</td>
<td width="618">�@</td>
<td width="36">�@</td>
<td width="20">�@</td>
</tr>
<tr> 
<td rowspan=2 width="11">�@</td>
<td class=bg1 colspan=2 height="73" 
style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
<table width='545' align=left cellspacing=1 border=1 cellpadding=0 bgcolor='#FFFFFF' bordercolor="#000000">
<tr bgcolor="#ffffff"> 
<td align=center width="32" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">ID</font></div>
</td>
<td align=center width="65" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">�m �W</font></div>
</td>
<td align=center width="48" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">�� ��</font></div>
</td>
<td align=center width="88" height="21" bgcolor="#FFFFFF"> <font size="2">����v</font> </td>
<td align=center width="90" height="21" bgcolor="#FFFFFF"> <font size="2">�벼</font> 
</td>
<td align=center width="241" height="21" bgcolor="#FFFFFF"> 
<div align="center"><font size="2">�����ާ@</font></div>
</td>
</tr>
<% 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("hg_connstr")
conn.open connstr
sql="SELECT * FROM �Τ� where ����='������' and ����<>'�x��' order by grade DESC"
Set Rs=conn.Execute(sql)
do while not rs.eof and not rs.bof
piao=rs("piao")
%>
<tr bgcolor="#C4DEFF"> 
<td width="32" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"><%=rs("id")%></font></div>
</td>
<td width="65" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#ff0000"><%=rs("�m�W")%></font></div>
</td>
<td width="48" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"><%=rs("grade")%></font></div>
</td>
<td width="88" bgcolor="#FFFFFF" height="2"> <font size="2" color="#000000"><img src=bar.gif width=<%=piao%> height=10><%=piao%>��</font> 
</td>
<td width="95" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"><a href=adddata.asp?id=<%=rs("id")%>>�����</a>|<a href=adddata2.asp?id=<%=rs("id")%>>��ĳ��</a></font></div>
</td>
<td width="241" bgcolor="#FFFFFF" height="2"> 
<div align="center"><font size="2" color="#000000"> 
<%if rs("grade")<=grade and rs("����")<>"�x��" then%>
<a href=admin1.asp?id=<%=rs("id")%>&a=c&sf=<%=rs("grade")+1%>>�ɯ�</a> 
| <a href=admin1.asp?id=<%=rs("id")%>&a=c&sf=<%=rs("grade")-1%>>����</a> 
| <a href=admin1.asp?id=<%=rs("id")%>&a=b>�}��</a> | 
<%end if%>
<a href=admin2.asp?id=<%=rs("id")%> title="�Ω�Q��������^�����������I">���^����</a></font></div>
</td>
</tr>
<%
rs.movenext
loop
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
</table>
<p><font size="2"> <br>
<br>
</font></p>
<p>�@</p>
<p>�@</p>
</td>
<td valign=top width="20" height="73">�@</td>
</tr>
<tr> 
<td class=bg1 colspan=2 
style="PADDING-LEFT: 15px; PADDING-RIGHT: 15px" valign=top> 
<form method=POST action='admin1.asp?a=a'>
<p><font size="2" color="#FFFFFF">�ۦ��G</font><font size="2"> 
<input name=name2 type=text size=10>
<input value="�T�w" type=submit name="submit2">
</font> </p>
</form>
<p>
<font color="#FF0000">&nbsp;����벼�G�������A��誺����v�[�@�A��Ϲﲼ�A��誺����v��@<br>
<font color="#FF0000">&nbsp;�����|�w���ˬd����v���t�ƪ��޲z���A�ä��H�B�@<br><br>
<span class="unnamed1"><font size="2" color="#FFFFFF">��������������6--10�šA���Ө��������C�֦����P���v�O�I�䤤�G<br>
<br>
ĵ�i�P�I���ޡ]��H�^�����v�O<br>
<br>
<font color="#FF0000">&nbsp;6��</font> �H�W�i�H�e��<br>
<font color="#FF0000">&nbsp;7��</font> �H�W�i�H�o��<br>
<font color="#FF0000">&nbsp;8��</font> �H�W�i�H��H���c�A24�p�ɤ��i�i����<br>
<font color="#FF0000">&nbsp;9��</font> �H�W�i�M�ĩM��ǤH�i��٭�<br>
<font color="#FF0000">10��</font> �i��ǤH�i��٭��i�޲z�ؿ��A�i�碌�������H�ɭ���<br>
<font color="#FF0000">�x��</font> �i�H�}�����������H</font></span><font color="#FFFFFF">, 
�i�H��O�H�ɨ�Q�šC</font></p>
</td>
<td width="20">�@</td>
</tr>
<tr> 
<td width="11">�@</td>
<td colspan=2>�@</td>
<td width="20">�@</td>
</tr>
</tbody> 
</table>
<p>�@</p>
</BODY></HTML>