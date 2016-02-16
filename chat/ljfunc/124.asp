<%
'購買6級會員
function bbbbb()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩,會員等級,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<1 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=1-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]分,再操作！');}</script>"
	Response.End
end if
if rs("銀兩")<6000000000  then
Response.Write "<script language=JavaScript>{alert('你的銀兩不夠，需要銀兩60億,才能購買！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 會員等級=6,銀兩=銀兩-6000000000,操作時間=now() where id="&info(9)
bbbbb=info(0)& "成功購買6級會員,真是開心啊!"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>