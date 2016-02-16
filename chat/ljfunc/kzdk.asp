<%
'購買2級會員
function kzdk()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 會員等級,銀兩,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
if rs("銀兩")<300000000000  then
Response.Write "<script language=JavaScript>{alert('你金錢不足夠，需要金錢300億,才能購買！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 會員等級=2,銀兩=銀兩-300000000000,操作時間=now() where id="&info(9)
kzdk=info(0)& " 多謝您購買2級會員^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>