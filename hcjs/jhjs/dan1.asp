<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
MsgBox "請進入聊天室再進入當鋪操作！！"
window.close()
</script>
<%response.end
end if
%>
<%
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhjs")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
id=lcase(trim(request("id")))
sl=int(abs(Request.form("sl")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
if InStr(sl,"or")<>0 or InStr(sl,"=")<>0 or InStr(sl,"`")<>0 or InStr(sl,"'")<>0 or InStr(sl," ")<>0 or InStr(sl," ")<>0 or InStr(sl,"'")<>0 or InStr(sl,chr(34))<>0 or InStr(sl,"\")<>0 or InStr(sl,",")<>0 or InStr(sl,"<")<>0 or InStr(sl,">")<>0 then Response.Redirect "../../error.asp?id=54"
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
rs.close
rs.open "SELECT 銀兩,數量,id FROM 物品 where ID=" & id & " and 類型<>'卡片' and 裝備=false and 擁有者="&info(9),conn
if rs.eof and rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有該物品或者該物品正裝備中！');window.close();</script>"
	response.end
else
sl1=rs("數量")
idd=rs("id")
if sl1>=sl then
conn.execute "delete * from 物品 where id="&idd
yin=int(rs("銀兩")/10)
conn.execute "update 用戶 set 銀兩=銀兩+" & yin & ",操作時間=now() where id="&info(9)
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：變賣不成功，你沒有那麼多物品可當！');parent.history.go(-1);</script>"
	response.end

end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "dan.asp"
end if
%>