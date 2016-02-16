<%
'購買法力
function kzdd()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,銀兩,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("銀兩")<500000000  then
Response.Write "<script language=JavaScript>{alert('你金錢不足夠，需要金錢5億,才能購買！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 資質=資質+1000000,銀兩=銀兩-500000000,操作時間=now() where id="&info(9)
kzdd=info(0)& " 多謝您購買法力100萬^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>