<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0

%>
<!--#include file="data2.asp"-->
<%
info=Session("info")
if info(0)="" then
response.redirect"warning.htm"
else
IF Request.ServerVariables("ReQuest_Method") = "POST" THEN
sheepname=request("nick")
if sheepname="" or len(sheepname)="" then

%>
<script language="Vbscript">
msgbox"�d���W�r��g�����T�C",0,"����"
history.back
</script>
<%
else
rs.open"select �ɶ�,����,��O,����,����id,���s,�g��,�ޯ�,����,����I,�ѧO from �d�����A where �W�r='"&sheepname&"' and user='"&info(0)&"' and ��O>0",conn,1,1
if rs.bof then
conn.execute "Delete * From �d�����A Where user='" &info(0)& "'"
conn.execute "Delete * From �ޯ�C�� Where �֦���='" &info(0)& "'"
conn.execute "Delete * From �D��C�� Where �֦���='" &info(0)& "'"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z���d���w�g���F�αz���O�o���d�����D�H�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
else
if date()-rs("�ɶ�")>=1 then
sql="update �d�����A set ����I=����I+4,�ɶ�=date() where user='"&info(0)&"'"
conn.Execute(sql)
rs.close
rs.open"select ��O,����,����,����id,�ޯ�,���s,�g��,����,����I,�ѧO from �d�����A where �W�r='"&sheepname&"' and user='"&info(0)&"' and ��O>0",conn,1,1
end if
power=rs("��O")
at=rs("����")
id=rs("����")
fy=rs("���s")
jy=rs("�g��")
zc=rs("����")
ap=rs("����I")
checker=rs("�ѧO")
jn=rs("�ޯ�")


%>

<html>

<head>
<title>�d���|���Ҧ�</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<meta http-equiv="Content-Type" content="text/html; charset=big5">
<link href="../../chat/ccss2.css" rel="stylesheet" type="text/css">
</head>
<body bgcolor="#FFFFFF" background="../../ff.jpg" text="#000000" leftmargin="0" topmargin="0">
<table border="1" width="648" cellspacing="1" cellpadding="0" align="center" bordercolor="#000000">
<tr>
<td width="642" align="center"><font color="#0000FF">�|���d���Ҧ� </font></td>
</tr>
<tr>
<td width="642">
<table border="0" width="100%" cellspacing="1" cellpadding="0"
height="20">
<tr>
<td class="p3" width="8%"><img border="0" src="image/<%=id%>.gif" ></td>
          <td class="p3" width="29%"><font color="#000000">�W�r�G<%=sheepname%> 
            --<%=jn%></font></td>
<td class="p3" colspan="2" width="63%"> <font color="#000000">�i<a href="change.asp">�󴫦W�r</a>�j�i<a href="showstunt.asp" target="_blank">�d���d���ޯ�</a>�j�i<a href="indexsheep.asp" target="_blank">���d���ө���</a>�j�i<a href="javascript:this.location.reload()">��s����</a>�j</font></td>
</tr>
</table>
<table border="0" width="103%" cellspacing="1" cellpadding="0"
height="192">
<tr>
          <td class="p2" width="25%" height="18"><font color="#000000">�� �O�G</font></td>
<td class="p2" width="25%" height="18"><font color="#000000"><%=power%>
</font></td>
          <td class="p2" width="25%" height="18"><font color="#000000">����I�G</font></td>
<td class="p2" width="25%" height="18"> <font color="#000000"><%=ap-checker%></font></td>
</tr>
<tr>
          <td class="p3" width="25%" height="18"><font color="#000000">�� ���G</font></td>
<td class="p3" width="25%" height="18"><font color="#000000"><%=at%></font></td>
          <td class="p3" width="25%" height="18"><font color="#000000">�g ��G</font></td>
<td class="p3" width="25%" height="18"><font color="#000000"><%=jy%></font></td>
</tr>
<tr>
          <td class="p2" width="25%" height="18"><font color="#000000">�� �s�G</font></td>
<td class="p2" width="25%" height="18"><font color="#000000"><%=fy%></font></td>
          <td class="p2" width="25%" height="18"><font color="#000000">�� �ۡG</font></td>
<td class="p2" width="25%" height="18"><font color="#000000"><%=zc%></font></td>
</tr>
<tr>
          <td class="p1" width="100%" colspan="4" height="17"><font color="#000000">�d���|���`�N�ƶ��G</font></td>
</tr>
<tr>
<td class="p3" width="100%" colspan="4" height="71">
<p>-�ϥιD��O���A�b�d���ө��ʶR���W�[�U���ݩʩΰ_<font color="#FF3333">�S�O�\��</font>�D�㪺�ϥΡC<br>
              -��m�i�W�[�d����<font color="#FF0000">�g��ȡB�����O�Ψ��s�O</font>�C<br>
              �����I�ܬ�0��A�N����i��|���A�C�Ѧ�<font color="#FF3333">4�I����I</font>�֥[�I<br>
              �_�I�N�|�����|�ϥ��d�����U���ݩʤj�T�����A���]�����|��U���ݺجƦܨϥ��d�����`�C<br>
              �𮧯��_�d������O�ȡA�C�Ѧ�<font color="#FF3333">4������I��</font>�A�Χ���N���i���d���i��|���A���൥�U�@�ѡC<br>
              �d���i�H�b��ѫǶi��<font color="#FF3333">����</font>���ΰ_�S�ġA�o�˰��i�@�ӱj�j���d���O�����i�֪����C</p>
</td>
</tr>
<tr>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='item.asp'" type="button" value="�ϥιD��"
name="B1" style="font-family: �s�ө���; font-size: 12px">
<br>
</td>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='test.asp?name=<%=request("nick")%>'" type="button" value="��m"
name="B13" style="font-family: �s�ө���; font-size: 12px">
</td>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='adv.asp?name=<%=request("nick")%>'" type="button" value="�_�I"
name="B132" style="font-family: �s�ө���; font-size: 12px">
</td>
<td class="p2" width="25%" height="38" align="center">
<input onClick="javascript:window.document.location.href='relax.asp?name=<%=request("nick")%>'" type="button" value="��"
name="B133" style="font-family: �s�ө���; font-size: 12px">
</td>
</tr>
</table>
</td>
</tr>
</table>
<%
end if
rs.close
conn.close
%>
</body>
</html>
<%
end if
end if
end if
set rs=nothing
set conn=nothing
%>