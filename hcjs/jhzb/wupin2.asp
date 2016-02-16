<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
name=info(0)
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/jhzb")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if

id=lcase(trim(request("ID")))
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or inStr(id,"&")<>0 or inStr(id,"&&")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
active=lcase(trim(request("active")))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 類型,等級,攻擊,防御,裝備 FROM 物品 WHERE 擁有者=" & info(9) & " and id=" & id,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你沒有該物品！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
leixing=rs("類型")
dj=rs("等級")

zbgj=rs("攻擊")
zbfy=rs("防御")
if  rs("裝備")=true then

rs.close
	rs.open "SELECT 等級 FROM 用戶 WHERE id="&info(9),conn
if dj>rs("等級") then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：該裝備需要"&dj&"級纔能裝備！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
select case leixing
case "頭盔"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='頭盔' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a1=0,b1=0 where id=" & info(9)
conn.execute "update 物品 set 裝備=false where id=" & id

case "左手"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='左手' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a2=0,b2=0 where id=" & info(9)
conn.execute "update 物品 set 裝備=false where id=" & id

case "右手"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='右手' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a3=0,b3=0 where id=" & info(9)
conn.execute "update 物品 set 裝備=false where id=" & id

case "雙腳"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='雙腳' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a4=0,b4=0 where id=" & info(9)
conn.execute "update 物品 set 裝備=false where id=" & id

case "裝飾"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='裝飾' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a5=0,b5=0 where id=" & info(9)
conn.execute "update 物品 set 裝備=false where id=" & id

case "盔甲"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='盔甲' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a6=0,b6=0 where id=" & info(9)
conn.execute "update 物品 set 裝備=false where id=" & id

case else
    rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：操作錯誤！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end select
else
select case leixing
case "頭盔"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='頭盔' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a1="&zbgj&",b1="&zbfy&" where id=" & info(9)
conn.execute "update 物品 set 裝備=true where id=" & id

case "左手"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='左手' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a2="&zbgj&",b2="&zbfy&" where id=" & info(9)
conn.execute "update 物品 set 裝備=true where id=" & id

case "右手"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='右手' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a3="&zbgj&",b3="&zbfy&" where id=" & info(9)
conn.execute "update 物品 set 裝備=true where id=" & id

case "雙腳"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='雙腳' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a4="&zbgj&",b4="&zbfy&" where id=" & info(9)
conn.execute "update 物品 set 裝備=true where id=" & id

case "裝飾"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='裝飾' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a5="&zbgj&",b5="&zbfy&" where id=" & info(9)
conn.execute "update 物品 set 裝備=true where id=" & id

case "盔甲"

rs.close
rs.open "SELECT id FROM 物品 WHERE 擁有者=" & info(9) & " and 類型='盔甲' and 裝備=true",conn
if not rs.eof and not rs.bof then
idd=rs("id")
conn.execute "update 物品 set 裝備=false where id=" & idd
end if
conn.execute "update 用戶 set a6="&zbgj&",b6="&zbfy&" where id=" & info(9)
conn.execute "update 物品 set 裝備=true where id=" & id

case else
    rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：操作錯誤！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end select
end if
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>location.href = 'head.asp';</script>"
	response.end
%>