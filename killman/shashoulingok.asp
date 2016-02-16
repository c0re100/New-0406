<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=request("name")
myname=info(0)
if info(0)<>"江湖總站" then Response.Redirect "../error.asp?id=439"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="delete * from 殺手 where 姓名1='" & name & "'"
conn.execute sql
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('已成功刪除"&name&"的殺人雇傭記錄');location.href = 'shashoulist.asp';}</script>"
else
Response.Write "<script language=JavaScript>{alert('你沒資格操作');location.href = 'shashoulist.asp';}</script>"
%>