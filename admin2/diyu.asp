<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
username=info(0)
if Instr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," "&LCase(info(0))&" ")=0 then
	Response.Write "<script Language=Javascript>alert('�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI');window.close();</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select ����,��O�[,��O from �Τ� where id="&info(9),conn
tlj=(rs("����")*1500+5260)+rs("��O�[")+1000
sm=rs("��O")
%>

<head>
<title>�a��</title>
</head>

<body bgcolor="#996699" text="#FFFFFF" oncontextmenu=self.event.returnValue=false LEFTMARGIN="0" TOPMARGIN="0">
<table width="493" border="0" cellspacing="0" cellpadding="0" align="CENTER" height="267"> 
<tr> <td height="267" width="508"> <table width="458" border="0" cellspacing="0" cellpadding="0" height="318"> 
<tr> <td height="298" width="1" ></td><td  background="../images/haidian_0928_2.jpg" width="486" align="center" height="298"><br> 
<div id="Layer2" style="z-index: 2; left: 222px; width: 267; position: absolute; top: 316px; height: 80"> 
<table cellspacing="0" border="0" width="251" bordercolorlight="#cccc99" bordercolordark="#FFFFFF" height="15"> 
<% if sm>=tlj then
response.write "�A�ثe�����A�ܦn�ڡA���ݭn�A�ҤF"
response.end
else%> <form method=POST action=diyu2.asp  onSubmit="this.ok.disabled=true"> <tr> 
<td height="1" width="247"> <div align="left">�A���ͩR�O��<%=sm%>, <%
if sm>=10000 then response.write "���ݭn�ҤF�I�A�ҴN�|����10000�F":p=1
if sm<10000 and sm>0 then response.write "���@�ŷҤ��I":p=2
if sm<=0 and sm>-5000 then response.write "���G�ŷҤ�":p=3
if sm<=-5000 then response.write "�ͩR���M�A�T���u���I":p=4
%> </div><tr> <td width="247"> <div align="left">�v����k�G <input name=yilao readonly size=12 value="<%
if p=1 then response.write "�^�h�a,���w��A"
if p=2 then response.write "�@�ŷҤ�"
if p=3 then response.write "�G�ŷҤ�"
if p=4 then response.write "�T���u��"
%>"> <input name=ok type=submit value=�T�w> &nbsp; </div></td></tr> </form><%
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end if%> </table></div><br> </td><td height="298" width="49"> </div> </td></tr> 
</table></td></tr> </table>
</body>
<html>