<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<11 then Response.Redirect "../error.asp?id=900"
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
conn.execute "update �Τ� set �ާ@�ɶ�=now(),�_��=now(),���B�ɶ�=now(),�޳N��=0"
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('�{�b�w�N�Ҧ��ާ@�ɶ����~�t���Τᰵ�F�վ�I');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>
