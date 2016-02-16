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
rs.open "SELECT id,计秖 FROM 珇 where 局Τ=" & info(9) & " and 计秖>0 and 珇='纒膟'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('⊿Τ赣贺膟!');location.href = 'javascript:history.go(-1)';}</script>"
	response.end
end if
id=rs("id")
sl=rs("计秖")
sl1=sl*5
conn.execute "update 珇 set 计秖=计秖-"&sl&" where id="&id
conn.execute "update ノめ set allvalue=allvalue+"&sl1&" where id="&info(9)
rs.close
set rs=nothing	
conn.close
set conn=nothing
Response.Write "<script language=JavaScript>{alert('Τ赣贺膟"&sl&",传竒喷"&sl1&"翴');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
%>
