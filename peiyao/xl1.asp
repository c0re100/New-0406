<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
name=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "SELECT id FROM 珇 WHERE 局Τ=" & info(9) & " and 计秖>=5 and 珇=''",conn
if rs.eof or rs.bof then 
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボぃ镑5ぃ巨');</script>"
	response.end
end if
iceid=rs("id")
conn.Execute "update 珇 set 计秖=计秖-5 WHERE id="&iceid
rs.close
rs.open ="select id from 絤珇 where 珇='礥ㄠ皊' and 局Τ="&info(9)
If Rs.Bof OR Rs.Eof then
'	conn.execute "insert into 絤珇(珇,局Τ,计秖,,糤计) values ('礥ㄠ皊',"&info(9)&",1,'ず',1500)"
	conn.execute "insert into 絤珇(珇,局Τ,计秖,,糤计) values ('礥ㄠ皊',"&info(9)&",1,'ず',1800)"
else
	id=rs("id")
	conn.execute "update 絤珇 set 糤计=1500,计秖=计秖+1 where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
%>
<script Language="Javascript">
alert("眤礥ㄠ皊絤ЧΘ")
parent.cz1.location.reload();
parent.ig.location.reload();
</script>
