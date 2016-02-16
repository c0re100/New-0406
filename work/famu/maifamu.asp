<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
my=info(0)
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"work/famu")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if 
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 職業 from 用戶 where id="&info(9),conn
if trim(rs("職業"))="" then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：您的職業有錯，所以您不能在這裡砍伐樹木！！');</script>"
	response.end
end if
if Session("minets")=1 then
	Session("minets")=Session("minets")-1
end if
if Session("minets")<=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你的木頭太少，賣不了幾個錢？？？！');</script>"
	response.end
end if
conn.execute "update 用戶 set 銀兩=銀兩+"& Session("minets")*800 &" where id="&info(9)
Response.Write "<script Language=Javascript>alert('您采木頭："& session("minets") &"塊,賣了"& session("minets")*800 &"兩白銀！');</script>"
Session("minets")=0
Session("minesj")=true
rs.close
set rs=nothing	
conn.close
set conn=nothing
%>
<script Language="Javascript">
parent.ig.location.reload();
</script>
