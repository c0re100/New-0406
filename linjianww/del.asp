<%@ LANGUAGE="VBScript" %>
<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
if info(2)=10 then
Response.Write "<script Language=Javascript>alert('�A�S���o���v���I');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
set rs=server.createobject("adodb.recordset")
sql="delete from ���~�R where id="&request("id")
rs.open sql,conn,1,1
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','���~�R��')"
conn.Execute(sql)
set rs=nothing
conn.close
Response.Redirect "../ok.asp?id=700"
%>
 