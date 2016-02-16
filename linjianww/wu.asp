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
if info(0)<>"打︽" then Response.Redirect "manerr.asp?id=255"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="update ユカ初 set 计秖=计秖+20 where 计秖>0 and 摸<>'猭腳' and 摸<>'猭竟' and 摸<>'' and よΑ='玂繧耫'"
Set Rs=conn.Execute(sql)
'sql="update 珇 set 计秖=计秖-20 where 计秖>0 and 摸='猭腳'"
'Set Rs=conn.Execute(sql)
'sql="update 珇 set 计秖=计秖-20 where 计秖>0 and 摸<>''"
'Set Rs=conn.Execute(sql)
'sql="update 珇 set 计秖=计秖+8 where 计秖>0 and 摸=''"
'Set Rs=conn.Execute(sql)
Set rs=nothing
Set conn=nothing
Response.Write "<script Language='Javascript'>"
Response.Write "alert('瞷盢┮Τ珇50');"
Response.Write "history.go(-1);"
Response.Write "</script>"
Response.End
%>
