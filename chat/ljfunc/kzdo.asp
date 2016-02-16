<%
'購買6級官府
function kzdo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select grade,銀兩,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
if rs("銀兩")<3000000000  then
Response.Write "<script language=JavaScript>{alert('你金錢不足夠，需要金錢30億,才能購買！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set grade=6,銀兩=銀兩-3000000000,操作時間=now() where id="&info(9)
kzdo=info(0)& " 多謝您購買6級官府^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>