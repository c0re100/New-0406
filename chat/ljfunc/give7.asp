<%
'領取防御
function give7()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 防御加,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("n",rs("操作時間"),now())
if sj<60 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('你距離拿防御的時間還有["& ss &"]分！請再等等啦！');}</script>"
	Response.End
end if
conn.execute "update 用戶 set 防御加=防御加+90000000,操作時間=now() where id="&info(9)
give7=info(0)& " 也存了點差不多時候了，已經拿了防御上限9000000"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function
%>