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
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT 单 FROM ノめ WHERE id=" & info(9),conn
if rs("单")<105 then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボ单ぃ镑汾硑猌竟');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "SELECT id FROM 珇 WHERE 局Τ=" & info(9) & " and 计秖>=260 and 珇=''",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボぃ镑260ぃ巨');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id1=rs("id")
rs.close
rs.open "SELECT id FROM 珇 WHERE 局Τ=" & info(9) & " and 计秖>=220 and 珇='膓ホ'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボ膓ホぃ镑220ぃ巨');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id2=rs("id")
rs.close
rs.open "SELECT id FROM 珇 WHERE 局Τ=" & info(9) & " and 计秖>=60 and 珇='屡腳ホ'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボ屡腳ホぃ镑60ぃ巨');</script>"
	response.end
end if
id3=rs("id")
rs.close
rs.open "SELECT id FROM 珇 WHERE 局Τ=" & info(9) & " and 计秖>=220 and 珇='堵ホ'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボ堵ホぃ镑220ぃ巨');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id4=rs("id")
rs.close
rs.open "SELECT id FROM 珇 WHERE 局Τ=" & info(9) & " and 计秖>=220 and 珇='垂'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボ垂ぃ镑220ぃ巨');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id5=rs("id")
rs.close
rs.open "SELECT id FROM 珇 WHERE 局Τ=" & info(9) & " and 计秖>=220 and 珇='Χホ'",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボΧホぃ镑220ぃ巨');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
id6=rs("id")
conn.Execute "update 珇 set 计秖=计秖-260 WHERE id="&id1
conn.Execute "update 珇 set 计秖=计秖-220 WHERE id="&id2
conn.Execute "update 珇 set 计秖=计秖-60 WHERE id="&id3
conn.Execute "update 珇 set 计秖=计秖-220 WHERE id="&id4
conn.Execute "update 珇 set 计秖=计秖-220 WHERE id="&id5
conn.Execute "update 珇 set 计秖=计秖-220 WHERE id="&id6
rs.close
rs.open "select id from 珇 where 珇='惦纒糃' and 局Τ="&info(9),conn
If Rs.Bof OR Rs.Eof then

conn.execute("insert into 珇(珇,局Τ,摸,ň眘,ず,砰,绊㏕,ю阑,弧,sm,计秖,单) values('惦纒糃',"&info(9)&",'も',0,0,0,2200000,600000,20032011,20032011,1,105)")

else
	id=rs("id")
	conn.execute "update 珇 set 计秖=计秖+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("惦纒糃汾硑ЧΘ")
history.back()
</script>
