<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="Delete * From �b�u���  "
Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('�{�b�w�N�Ȩ�䧽�i��F�ץ��I');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>
