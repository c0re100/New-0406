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
if info(2)<10   then Response.Redirect "../error.asp?id=900"
wupinname=trim(Request.Form("wupinname"))
wupinlx=trim(Request.Form("wupinlx"))
wupinnl=abs(trim(Request.Form("wupinnl")))
wupintl=abs(trim(Request.Form("wupintl")))
wupingj=abs(trim(Request.Form("wupingj")))
wupinsm=trim(Request.Form("wupinsm"))
wupinfy=abs(trim(Request.Form("wupinfy")))
wupinjg=abs(trim(Request.Form("wupinjg")))
wupindj=abs(Request.Form("wupindj"))
wupinyl=abs(Request.Form("wupinyl"))
id=trim(Request.Form("wupinID"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.Execute "update ���~�R set ���~�W='"& wupinname &"',����='"& wupinlx &"',���O="& wupinnl &",��O="& wupintl &",����="& wupingj &",����='"& wupinsm &"',���s="& wupinfy &",��T��='"& wupinjg &"',����="& wupindj &",�Ȩ�="& wupinyl &" where id="&ID
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�L���ק�ާ@')"
conn.Execute(sql)
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../ok.asp?id=700"
%> 