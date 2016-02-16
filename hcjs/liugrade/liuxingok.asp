<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/liugrade")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
id=lcase(trim(request("id")))
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or inStr(id,"&")<>0 or inStr(id,"&&")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 物品 where 物品名='流星' and 數量>0 and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('升級武器需要流星！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
idd=rs("id")
conn.execute "update 物品 set 數量=數量-1 where id="&idd
rs.close
rs.open "select 類型,等級,攻擊,防御 from 物品 where id="&id&" and 數量>0 and 裝備=false and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器正裝備著呢，若要升級請解除裝備！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
dj=rs("等級")
gj=rs("攻擊")
fy=rs("防御")
randomize timer
r=int(rnd*dj)
if r<8 then
if rs("類型")="右手" then
conn.execute "update 物品 set 等級=等級+5,攻擊=攻擊+500 where id="&id
end if
if rs("類型")="左手"  then
conn.execute "update 物品 set 等級=等級+5,防御=防御+500 where id="&id
end if
if  rs("類型")="盔甲" then
conn.execute "update 物品 set 等級=等級+5,防御=防御+200 where id="&id
end if
if  rs("類型")="頭盔" then
conn.execute "update 物品 set 等級=等級+5,防御=防御+400 where id="&id
end if
if rs("類型")="雙腳" then
conn.execute "update 物品 set 等級=等級+5,攻擊=攻擊+100,防御=防御+100 where id="&id
end if
if rs("類型")="裝飾" then
conn.execute "update 物品 set 等級=等級+5,防御=防御+100 where id="&id
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('升級武器成功');location.href = 'javascript:history.go(-1)';}</script>"
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('升級武器失敗');location.href = 'javascript:history.go(-1)';}</script>"
end if%>