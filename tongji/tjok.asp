<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<6 then Response.Redirect "../error.asp?id=461" 
if trim(request.form("tjname"))="" or trim(request.form("yuanyin"))="" or trim(Request.Form("fuyan"))="" then %> 
<script language=vbscript> 
MsgBox "���~�G�q�r�H�ǡB�q�r��]�B�ШD����������g�I" 
location.href = "javascript:history.back()" 
</script> 
<%end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

'rs.open "select ���� from �Τ� where �m�W='" & request.form("tjname") & "'"
	sql="SELECT * FROM �Τ� WHERE �m�W='" & request.form("tjname") & "'"
	Set Rs=conn.Execute(sql)
	if rs.eof or rs.bof then

	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=javascript>alert('�A�q�r���ëD����H���A�ХJ���ˬd�z��J���q�r�H�Ǫ���g���T!');history.back();</script>"
	response.end
end if%>
<%
tjname=request.form("tjname")
yuanyin=request.form("yuanyin")
fuyan=request.form("fuyan")
fabu=info(0)
%>
<!--#include file="data1.asp"--><%
'rs.close
'rs.open "select �q�r�H�� from �q�r�O where �q�r�H��='" & request.form("tjname") & "'"
sql="SELECT * FROM �q�r�O WHERE �q�r�H��='" & request.form("tjname") & "'"
	Set Rs=conn.Execute(sql)
if rs.eof or rs.bof then
                sql="insert into �q�r�O(�q�r�H��,�ШD�H��,�q�r��],�ШD����,�ШD�ɶ�,�����f��,��Ǯɶ�,��ǤH��) values ('" & tjname & "','" & fabu & "','" & yuanyin & "','" & fuyan & "',now(),'�ݼf','����','����')"
                        conn.execute sql
tj="<font color=red>���[" & fabu & "]���\����q�r�ШD�A�q�r�H�ǡG[" & tjname & "]�A�q�r��]�G[" & yuanyin & "]�A���A�G���ݯ����f��I</font>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
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
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�q�r�ШD�j</font>"&tj
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
<script language=vbscript> 
MsgBox "�����G�z�w���\�o���q�r�ШD�A�Э@�ߵ��ݯ����f��I" 
location.href = "javascript:history.back()" 
</script>
<%else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript> 
MsgBox "���~�G�ӤH�Ǥw�g�b�ШD�W�椺�A�Фŭ��_��g�I" 
location.href = "javascript:history.back()" 
</script> 

<%end if%>