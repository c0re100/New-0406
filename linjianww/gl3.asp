<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5" %>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
nickname=info(0)
if info(2)<11 then Response.Redirect "../error.asp?id=900"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="delete * from �Τ� where CDate(lasttime)<now()-30 and �|������=0"
Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('30�ѱq���n�����Τ�w�g�����R���I');"
Response.Write "location.href = 'gl.asp';"
Response.Write "</script>"
Response.End
%>
<html></html> 