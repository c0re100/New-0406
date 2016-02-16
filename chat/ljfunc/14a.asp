<%'搶奪

function qiang(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('搶奪對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(2)>6 then
	Response.Write "<script language=JavaScript>{alert('官府人員不可進行此操作！');}</script>"
	Response.End
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(2)<>1 then
Response.Write "<script Language=Javascript>alert('要搶寶請去大廳！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 寶物 from 用戶 where id="&info(9),conn
if rs("寶物")="靈劍水晶石" then
		Response.Write "<script language=JavaScript>{alert('你自己有寶物還要搶人家的呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
'對方的判斷
rs.open "select 寶物,寶物修練 from 用戶 where 姓名='"&towho&"'",conn

if rs("寶物")="靈劍水晶石"  then
		conn.execute "update 用戶 set 寶物修練=0,寶物='靈劍水晶石' where id="&info(9)
		conn.execute "update 用戶 set 寶物修練=0,寶物='無' where 姓名='"&towho&"'"
	qiang=info(0) & "把"& towho &"的寶物:靈劍水晶石搶走，因得到此寶。江湖寶物需要進行修練才可以得到更多的東西！"
		
else
qiang=info(0) & "想搶"& towho &"的寶物,可是"& towho &"身上並沒有什麼寶物啊!"
		
end if
rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
end function

%>