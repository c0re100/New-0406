<%
function xiulian()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 寶物,寶物修練,等級,內力,操作時間 FROM 用戶 WHERE id="&info(9),conn
if rs("寶物")<>"靈劍水晶石" then
	Response.Write "<script language=JavaScript>{alert('你並沒有江湖寶物[靈劍水晶石]操作錯誤！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("寶物修練")>=int(rs("等級")/10)+1 then
	conn.execute "update 用戶 set 寶物='無',法力=法力+100,內力加=內力加+寶物修練,體力加=體力加+寶物修練,武功加=武功加+寶物修練,道德加=道德加+寶物修練,魅力加=魅力加+寶物修練,攻擊加=攻擊加+寶物修練,防御加=防御加+寶物修練,銀兩=銀兩+(200000*寶物修練) where id="&info(9)
	conn.execute "update 用戶 set 寶物修練=0,操作時間=now() where id="&info(9)
	xiulian=info(0) & "<bgsound src=wav/xl.wav loop=1>祝賀您,您把你的的江湖寶物<font color=red>靈劍水晶石</font>修練完成,所有上限加<font color=red>"& rs("寶物修練") &"</font>點,得現金：+"& 200000*rs("寶物修練") &"兩,法力+100點,寶物修練完成，自動消失了。。。。。。"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
if chatinfo(5)<>0 then
sj=DateDiff("s",rs("操作時間"),now())
if sj<60 then
	s=60-sj
	Response.Write "<script language=JavaScript>{alert('你還沒有修練完成，請等["& s &"]秒再進行操作！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
end if
if rs("內力")<150  then
	Response.Write "<script language=JavaScript>{alert('內力少於150，無法修練！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 寶物修練=寶物修練+1,內力=內力-150,操作時間=now() where id="&info(9)
xiulian=info(0) & "擁有江湖寶物靈劍水晶石進行修練，這是你第:"& rs("寶物修練")+1 & "次進行寶物修練。。"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>