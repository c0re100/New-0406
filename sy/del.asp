<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(0)<>"¦¿´ò¦æ"  then Response.Redirect "../error.asp?id=439"
id=request("id")

sql="delete * from ¥Ó­Þ  where id=" & id & ""
conn.execute sql

conn.close
set rs=nothing	
set conn=nothing
Response.Redirect "manage.asp"
%>