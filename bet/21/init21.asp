<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

sql="Select �y�O,�Ȩ� from �Τ� where id="&info(9)
set rs=conn.Execute(sql)
nowmeili=rs("�y�O")
if nowmeili<10 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "�j�L�A�z�ثe���y�O�C��10�F�A���l�W�ݤ��L�h�r�A�H��A�ӧa�I"
location.href = "javascript:history.back()"
</script>
<%
else
%>
<!--#include file="h_gamble.asp"-->
<%
Session("n_PokerBack")="deck_2.gif"
'Session("n_TotalMoney")=1000
Session("n_TotalMoney")=rs("�Ȩ�")
Session("n_Bet")=0
Session("bet_win")=0
Session("n_Init")=false '�Ĥ@���i�J��խ����Ч�
	
redim userPokers(1,0)
redim antiPokers(1,0)
		
Session("n_UserPokers")=userPokers '�Τ�P
Session("n_AntiPokers")=antiPokers '�q���P
Session("n_Result")=4
	
rs.close
set rs=nothing
			      		
		
Response.Redirect "gamble.asp"
%>
<%
end if
%>
<html></html>