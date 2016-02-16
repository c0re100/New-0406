<%
'購買水天師
function kzdv()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 職業,銀兩,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
if rs("銀兩")<1500000000  then
Response.Write "<script language=JavaScript>{alert('你金錢不足夠，需要金錢15億,才能購買！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 職業=水天師,銀兩=銀兩-1500000000,操作時間=now() where id="&info(9)
kzdv=info(0)& " 多謝您購買水天師^.^"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>