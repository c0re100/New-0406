<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
%>
<!--#include file="h_gamble.asp"-->
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"

Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")

dim bet
dim over
	
dim pages
pages="gamble.asp"
	
bet=trim(Request.Form("Bet"))
	
select case Request.Form("Action")
case "�}��"
over=0
call GameStart(bet)
case "�A�Ӥ@��"
over=0
call GameStart(bet)
case "�n�P"
over=0
call GiveUserPoker()
if CanAddPoker() then
GiveAntiPoker()
end if
case "�}�P"
if CanAddPoker() then
over=0
GiveAntiPoker()
else
''''''''''''''''''''''''
Session("n_Result")=Result()	
Session("n_Begin")=false
			
dim value	
over=0	
value=CLng(Session("n_Bet"))
Session("n_TotalMoney")=Session("n_TotalMoney")+value
'value=CCur(Session("n_Bet"))
select case CInt(Session("n_Result"))
case 1 'win
				
value=value
case 2 'lost
value=0-value
case 3
value=0
case else
value=0
end select
conn.execute("Update �Τ� set �Ȩ�=�Ȩ�+"&value&",�y�O=�y�O-1 where �m�W='"&info(0)&"'") 				
Session("n_TotalMoney")=Session("n_TotalMoney")+value		
Session("n_Bet")=0
end if
case "�����F"
pages="../jh.asp"
case "��^"
'�b�o�̪�l�Ƽƾ�
'	Session("n_Init")=false
'	Session("n_TotalMoney")=1000
'case "�����F"
'�h�X�a�C
pages="../jh.asp"
case else
call GameStart(bet)
end select

if over=0 then

Response.Redirect pages
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>