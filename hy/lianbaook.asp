<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 寶物,知質 from 用戶 where id="&info(9),conn
if rs("知質")<1000  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的資質低於1000煉不出什麼的！');window.close();}</script>"
Response.End
end if
if rs("寶物")<>"靈劍水晶石"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你沒有寶物靈劍水晶石啊！');window.close();}</script>"
Response.End
end if
zhizi=rs("知質")
conn.execute "update 用戶 set 法力=法力+"&zhizi&",知質=0,寶物='無' where id="&info(9)
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('煉寶成功，你的法力增長"&zhizi&"！');window.close();}</script>"%>