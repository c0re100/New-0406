<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
nickname=info(0)
if info(0)<>"�����" then Response.Redirect "manerr.asp?id=255"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update ������� set �ƶq=�ƶq+20 where �ƶq>0 and ����<>'�k�_' and ����<>'�k��' and ����<>'�d��' and �覡='�O�I�d'"
Set Rs=conn.Execute(sql)
'sql="update ���~ set �ƶq=�ƶq-20 where �ƶq>0 and ����='�k�_'"
'Set Rs=conn.Execute(sql)
'sql="update ���~ set �ƶq=�ƶq-20 where �ƶq>0 and ����<>'�d��'"
'Set Rs=conn.Execute(sql)
'sql="update ���~ set �ƶq=�ƶq+8 where �ƶq>0 and ����='�d��'"
'Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('�{�b�w�N�Ҧ����~�[�F50�ӡI');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>
