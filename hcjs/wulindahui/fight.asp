<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="data.asp"-->
<%
my=info(0)
types=Request.QueryString("typename")
id=Request.QueryString("id")
Select Case id
Case "1"
  classes="����"
Case "2" 
   classes="�]��"
Case "3"
   classes="����"
End Select
Select Case types
Case "gold"
  typeses="���]"
Case "silver" 
   typeses="�Ⱥ]"
Case "copper"
   typeses="�ɺ]"
End Select
%>
<%
sql="select �Z�\,���s,����,���O,����,��O,�ʧO from �Τ� where �m�W='" & my & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
	Response.Redirect "../../error.asp?id=454"
	conn.close
	response.end
end if
set rs1=server.createobject("adodb.recordset")
sql1="select ����,���s,���O,�Z�\ from "&types&" where id="&id
set rs1=conn1.execute(sql1)
if rs("�Z�\")>10000 and types="copper" then
%>
<script language=vbscript>  
  MsgBox "�A�p���F�`���Z�\�w�g���h��Ӵ��ɺ]�F�C"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("�Z�\")<=10000 and types="silver" then
%>
<script language=vbscript>  
  MsgBox "�A�p���t���Z�\�ٷQ�Ӵ��Ⱥ]�H�^�h�A�m�m�a�I"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("�Z�\")>20000 and types="silver" then
%>
<script language=vbscript>  
  MsgBox "�A�p���F�`���Z�\�w�g���h��Ӵ��Ⱥ]�F�C"  
  location.href = "meeting.asp"  
</script>
<%
rs.close
rs1.close
conn.close
conn1.close
response.end
end if
if rs("�Z�\")<20000 and types="gold" then
%>
<script language=vbscript>  
  MsgBox "�A�p���t���Z�\�ٷQ�Ӵ����]�H�^�h�A�m�m�a�I"  
  location.href = "meeting.asp"  
</script>
<%rs.close
rs1.close
conn.close
conn1.close
response.end
end if
set rs2=server.createobject("adodb.recordset")
sql2="select id from "&types&" where  �m�W='" & my & "'"
set rs2=conn1.execute(sql2)
if not rs2.eof then
if (rs2("id")<3 and id=3) or (rs2("id")<2 and id=2) then
%>
<script language=vbscript>
MsgBox "�A�w�g���o��o�ӧ󰪪��\�W�F�A�ӳg�ߤF�a�I"
location.href = "meeting.asp"
</script>
<%
rs.close
rs1.close
rs2.close
conn.close
conn1.close
response.end
end if
if (rs2("id")=3 and id=3) or (rs2("id")=2 and id=2) or (rs2("id")=1 and id=1) then
%>
<script language=vbscript>
MsgBox "�A�w�g���o�o�ӥ\�W�F�A�Q�M�ۤv�L���h��?"
location.href = "meeting.asp"
</script>
<%
rs.close
rs1.close
rs2.close
conn.close
conn1.close
response.end
end if
end if
if rs("���s")>rs1("���s") and rs("����")+rs("���O")>rs1("����")+rs1("���O") then
 wugong=rs("�Z�\")-rs1("�Z�\")
 defence=rs("���s")-rs1("���s")
 force=rs("����")-rs1("����")
 neili=rs("���O")-rs1("���O")
 sql1="update "&types&" set �m�W='"& my & "',�ʧO='"&rs("�ʧO")&"',����='"&rs("����")&"',����='"&rs("����")&"',�Z�\="&rs("�Z�\")&",���O="&rs("���O")&",��O="&rs("��O")&",����="&rs("����")&",���s="&rs("���s")&" where id="&id&""
conn1.execute(sql1)
sql="update �Τ� set ���O=���O+500,��O=��O-50,allvalue=allvalue+100, �Ȩ�=�Ȩ�+500000 where �m�W='"&my&"'"
conn.execute(sql)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=FFFDAF>�i���i�j</font><font color=#DFF0CF>����<font color=#DFF0CF>"& my &"</font>���]���\�A�n�W�F<font color=#DFF0CF>"& typeses & classes &"</font>���_�y!�������y�A���O5000�A�g��100�I�A�Ȥl500000�I�I"&"</font>" 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script language=vbscript>  
  MsgBox "���ߧA�A�A���]���\�I"  
  location.href = "meeting.asp"  
</script>
<%
else
sql="update �Τ� set  ���O=���O-50,��O=��O-50,�Ȩ�=�Ȩ�-500 where �m�W='"&my&"'"
conn.execute(sql)
%>
<script language=vbscript>  
  MsgBox "�藍�_�I�A���]���ѡI���G�Q���o��C�y�~���C"  
  location.href = "meeting.asp"
</script>
<%
end if
%>