<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
%>
<!--#include file="data1.asp"--><%
sql="select * from �Ы� where ��D='" & info(0) & "' or ��Q='" & info(0) & "'"
set rs=conn.execute(sql)
if rs.eof or rs.bof then
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�z�٨S���R�Фl�O�I');location.href = 'fangwu.asp';}</script>"
Response.End
end if
set rs=nothing
conn.close
set conn=nothing
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select * from �Τ� where id="&info(9)
set rs=conn.execute(sql)
%>
<html>
<head>
<style>
body{font-size:9pt;color:#000000;}
p{font-size:16;color:#000000;}
</style>
</head>
<body background="by.gif" bgproperties="fixed" bgcolor="#000000" vlink="#000000">
<center>
<%
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+100 where id="&info(9)
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('�A�ݹL�]�k�ѡA�@���p�ߴN�ܥX�Ӥ@�j����@�A�n���ְ�,�A���Ȩ�W�[100��I');location.href = 'moshu.asp';}</script>"
Response.End
%>
</body>
</html>
<html><script language="JavaScript">                                                                  </script></html>