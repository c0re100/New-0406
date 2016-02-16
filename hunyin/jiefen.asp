<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=request("name")
my=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT ﹎1 FROM るρ where ﹎1='" & name & "'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=481"
	response.end
end if
rs.close
rs.open "select 皌案,┦ from ノめ where ﹎='"&name&"'",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=427"
	response.end
else
if rs("皌案")="礚" then
sex=rs("┦")
rs.close
rs.open "SELECT ┦,皌案 FROM ノめ where id=" & info(9),conn
if rs("┦")=sex then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "../error.asp?id=428"
	response.end
end if			
if rs("皌案")="礚" then
	conn.execute "update ノめ set 皌案='" & name & "',挡盉丁=date() where id="&info(9)
	conn.execute "update ノめ set 皌案='" & my & "',挡盉丁=date() where ﹎='" & name & "'"
	conn.execute "delete * from るρ where ﹎1='" & name & "' or ﹎2='" & my & "'or ﹎1='" & my & "'or ﹎2='" & name & "'"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Redirect "yuelao.asp"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Redirect "../error.asp?id=429"
response.end
end if
else
conn.close
Response.Redirect "../error.asp?id=430"
response.end
end if
end if
%> 