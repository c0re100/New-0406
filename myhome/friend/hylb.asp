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


<!--#include file="data2.asp"-->
<%
pagecount=10

dqpage=cint(Request.QueryString("page"))
if dqpage=0 then
dqpage=1
end if

searchsql="select * from usertype where user='"&info(0)&"' and type='��' order by id desc"

rs.open searchsql,conn,1,1

if rs.bof then
%>

<script language="Vbscript">
msgbox "�S���z�ݭn���H���I",0,"����"
history.back
</script>

<%
Response.End
end if
temp1=rs.recordcount/pagecount
if temp1-int(rs.recordcount/pagecount)<>0 then
totalpage=int(temp1+1)
else
totalpage=temp1
end if
if dqpage>totalpage then
dqpage=1
end if

if dqpage<>1 then
temp2=(dqpage-1)*pagecount
rs.move temp2

end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<title>�ү���ϡV�ڪ��d���O</title>
<link rel="StyleSheet" href="../../new.css" title="Contemporary">
</head>

<body topmargin="0" leftmargin="0">

<div align="center">
<center>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="0" cellpadding="0">
<tr>
<td class=p1 width="476">��-�z��e����m-�L�ȴJ��---�n���W��----�n�ͦC��</td>
<td width="10"></td>
<td class=p1 width="290">��-��e�ɶ��G<%=date%><%=time%></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="776" cellspacing="0" cellpadding="0">
<tr>
<td width="100%"></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="589" cellspacing="1" cellpadding="0">
<tr>
<td class=p1 width="53"><a href="javascript:history.back(1)">��^</a></td>
</tr>
</table>
</center>
</div>
<div align="center">
<center>
<table border="0" width="470" cellspacing="0" cellpadding="0">
<tr>
<td width="176" valign="top">
<table border="0" width="474" cellspacing="1" cellpadding="0">
<tr>
<td class=p2 width="78">

<%if dqpage>1 then%>
<FORM METHOD=get ACTION="more_friend.asp">
<INPUT TYPE="hidden" name="page" value="1">
<INPUT TYPE="submit" value="����">
</FORM>

<%else%>
<INPUT TYPE="submit" value="����">
<%end if%>

</td>
<td class=p2 width="78">

<%if dqpage>1 then%>
<FORM METHOD=get ACTION="more_friend.asp">
<INPUT TYPE="hidden" name="page" value="<%=dqpage-1%>">
<INPUT TYPE="submit" value="�W��">
</FORM>


<%else%>
<INPUT TYPE="submit" value="�W��">
<%end if%>


</td>
<td class=p2 width="79">
<%if dqpage<totalpage then%>
<FORM METHOD=get ACTION="more_friend.asp">
<INPUT TYPE="hidden" name="page" value="<%=dqpage+1%>">
<INPUT TYPE="submit" value="�U��">
</FORM>


<%else%>
<INPUT TYPE="submit" value="�U��">
<%end if%>


</td>
<td class=p2 width="79">
<%if dqpage<totalpage then%>
<FORM METHOD=get ACTION="more_friend.asp">
<INPUT TYPE="hidden" name="page" value="<%=totalpage%>">
<INPUT TYPE="submit" value="����">
</FORM>


<%else%>
<INPUT TYPE="submit" value="����">


<%end if%>
</td>
<td class=p2 width="79">��e���G<%=dqpage%></td>
<td class=p2 width="79">�`���G<%=totalpage%></td>
</tr>
</table>
<table border="0" width="474" cellspacing="1" cellpadding="0">
<tr>
<td class=p2 width="300">�����m�W</td>
<td class=p2 width="174">�R���_�H</td>
</tr>
<%
count=0
do while not rs.eof and count<pagecount

%>
<tr>
<td class=p3 width="300"><a href='../guestbook/send.asp?name=<%=rs("name")%>'><%=rs("name")%></a> </td>
<td class=p3 width="174"><A HREF="deltype.asp?id=<%=rs("id")%>">�R��</A></td>
</tr>
<%
count=count+1
rs.movenext
loop
if count<pagecount then
for i=1 to pagecount-count

%>
<tr>
<td class=p3 width="300"> </td>
<td class=p3 width="174"> </td>
</tr>
<%
next
end if
rs.close
%>
</table>
<table border="0" width="476" cellspacing="1" cellpadding="0">
<tr>
<td class=p3 width="476">&nbsp;</td>
</tr>
<tr>
<td class=p3 width="476">&nbsp;</td>
</tr>
<tr>
<td class=p3 width="476"></td>
</tr>
<tr>
<td class=p3 width="476">&nbsp;</td>
</tr>
<tr>
<td class=p3 width="476">&nbsp;</td>
</tr>
<tr>
<td class=p3 width="476">&nbsp;</td>
</tr>
<tr>
<td class=p3 width="476">&nbsp;&nbsp;</td>
</tr>
</table>
</td>
</center>
</table>
</body>
<%conn.close%>

