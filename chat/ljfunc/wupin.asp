<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if chatbgcolor="" then chatbgcolor="008888"
%>
<html>
<head>
<title>�d�ݪ��~</title>
<script language="JavaScript">function s(list){parent.f2.document.af.sytemp.value=parent.f2.document.af.sytemp.value+list;parent.f2.document.af.sytemp.focus();}</script>
<style>
td{font-size:9pt;}
</style>
</head>
<body oncontextmenu=self.event.returnValue=false bgcolor="#660000" bgproperties="fixed" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" topmargin="0" leftmargin=0 background=bk.jpg>
<div align="center"><img src="ico/girl01.gif" width="24" height="41"><img src="ico/girl02.gif" width="21" height="38"> 
</div>
<table border="1" cellspacing="0" cellpadding="1" width="135" bordercolorlight="#000000" bordercolordark="#FFFFFF" align="center" >
<tr><td colspan=3 bgcolor=0090DF><div align="center"><a href="../hcjs/es/mybxg.asp" target="_blank"><img src="ico/03.gif" width="9" height="11" border="0">����O�I�d</a><img src="ico/03.gif" width="9" height="11" border="0"></div></td></tr>
<tr><td colspan=3 bgcolor=0090DF><div align="center"><a href="../hcjs/jhzb/eat.asp" target="_blank"><img src="ico/10.gif" width="9" height="9" border="0">�b�u�Y��</a><img src="ico/10.gif" width="9" height="9"></div></td></tr>
<tr><td colspan=3 bgcolor=0090DF><div align="center"><a href="#" onClick="window.open('../hcjs/jhzb/head.asp','','scrollbars=yes,resizable=yes,width=750,height=452')"><img src="ico/4.gif" width="11" height="11" border="0">�Z���˳�</a><img src="ico/4.gif" width="11" height="11"></div></td></tr>
<tr align=center bgcolor=ffcc00><td>�� �~ �W</td><td>�ƶq</td><td>�ۥ�</td></tr>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.execute "delete * from ���~ where �ƶq<=0"
rs.open "SELECT * FROM ���~ WHERE �֦���=" & info(9) & " and �˳�=false order by ����,�Ȩ� ",conn
do while not rs.bof and not rs.eof
lx=trim(rs("����"))
%>
<tr bgcolor="#EEFFEE"  onmouseout="this.bgColor='#EEFFEE';"onmouseover="this.bgColor='#DFEFFF';">
<td width="69">
<div align="center"><%=rs("���~�W")%> <img src="../hcjs/jhjs/001/<%=rs("sm")%>.gif" border="0" width="32" height="32"%></div>
</td>
<td width="26">
<div align="center"><%=rs("�ƶq")%> </div>
<div align="center"></div>
</td>
<td width="26">
<div align="center">
<%if trim(rs("����"))="�r��" then%>
<a href="javascript:s('/�U�r$ <%=rs("���~�W")%>')" title=��O:<%=rs("��O")%>���O:<%=rs("���O")%>�Ȩ�:<%=rs("�Ȩ�")%>><font color="#FF0000">�U�r</font></a>
<font color="#FFFFFF">
<%end if%>
<%if trim(rs("����"))="�t��" then%>
</font> <a href="javascript:s('/���Y$ <%=rs("���~�W")%>')" title=��O:<%=rs("��O")%>���O:<%=rs("���O")%>�Ȩ�:<%=rs("�Ȩ�")%>><font color="#FF0000">���Y</font></a>
<%end if%>
<%if trim(rs("����"))="�d��" then%>
<a href="javascript:s('/�ϥ�$ <%=rs("���~�W")%>')" title=����:<%=rs("����")%>����:<%=rs("�Ȩ�")%>��><font color="#FF0000">�ϥ�</font></a>
<a href="javascript:s('/�ذe$ <%=rs("���~�W")%>&1')" title=����:<%=rs("����")%>����:<%=rs("�Ȩ�")%>��><font color="#FF0000">�ذe</font></a>
<%end if%>
<%if trim(rs("����"))<>"�t��" and trim(rs("����"))<>"�r��" and trim(rs("����"))<>"�d��" then%>
<a href="javascript:s('/�ذe$ <%=rs("���~�W")%>&1')" title=����:<%=rs("����")%>���s:<%=rs("���s")%>��O:<%=rs("��O")%>���O:<%=rs("���O")%>��T��:<%=rs("��T��")%>�Ȩ�:<%=rs("�Ȩ�")%>����:<%=rs("����")%>><font color="#FF0000">�ذe</font></a>
<%end if%>
<a href="javascript:s('/���$ <%=rs("���~�W")%>&1')" title=����:<%=rs("����")%>���s:<%=rs("���s")%>��O:<%=rs("��O")%>���O:<%=rs("���O")%>�Ȩ�:<%=rs("�Ȩ�")%>����:<%=rs("����")%>><font color="#0000FF">���F</font></a>
</div>
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
</table></body></html>
