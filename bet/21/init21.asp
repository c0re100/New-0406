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

sql="Select 魅力,銀兩 from 用戶 where id="&info(9)
set rs=conn.Execute(sql)
nowmeili=rs("魅力")
if nowmeili<10 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language=vbscript>
MsgBox "大俠，您目前的魅力低於10了，面子上看不過去呀，以後再來吧！"
location.href = "javascript:history.back()"
</script>
<%
else
%>
<!--#include file="h_gamble.asp"-->
<%
Session("n_PokerBack")="deck_2.gif"
'Session("n_TotalMoney")=1000
Session("n_TotalMoney")=rs("銀兩")
Session("n_Bet")=0
Session("bet_win")=0
Session("n_Init")=false '第一次進入賭博頁面標志
	
redim userPokers(1,0)
redim antiPokers(1,0)
		
Session("n_UserPokers")=userPokers '用戶牌
Session("n_AntiPokers")=antiPokers '電腦牌
Session("n_Result")=4
	
rs.close
set rs=nothing
			      		
		
Response.Redirect "gamble.asp"
%>
<%
end if
%>
<html></html>