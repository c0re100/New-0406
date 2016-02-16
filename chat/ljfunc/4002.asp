<%
'尋找龍珠
function dddddd(fn1,to1,toname)
if toname="大家" or toname=info(0) or toname=Application("ljjh_automanname")  then
	Response.Write "<script language=JavaScript>{alert('尋龍珠對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<10 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要10級以上方可尋流星！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 法力,操作時間 FROM 用戶 WHERE id="&info(9),conn
fla=rs("法力")
sj=DateDiff("s",rs("操作時間"),now())
if sj<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=10-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒鐘,再操作！');}</script>"
	Response.End
end if
if fla<10000 then
Response.Write "<script language=JavaScript>{alert('你的法力不夠無法施展呀，至少也得10000點啊！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 法力=法力-100000,操作時間=now() where id="&info(9)
randomize()
rnd1=int(rnd()*4)
if rnd1<3 then
dddddd=info(0) & "好可惜，你尋遍了江湖各個角落也沒有找到什麼龍珠,"&info(0) & "損耗100000點法力!" 
else
rs.close
rs.open "select id FROM 物品 WHERE 物品名='龍珠' and 擁有者=" & info(9)
	If Rs.Bof OR Rs.Eof then
		conn.execute "insert into 物品(物品名,擁有者,類型,堅固度,攻擊,防御,內力,體力,數量,銀兩,說明) values ('龍珠',"&info(9)&",'物品',100,0,0,0,0,1,200000,1600)"
	else
id=rs("id")
		conn.execute "update 物品 set 銀兩=200000,數量=數量+1 where id="& id &" and 擁有者=" & info(9)
	end if
dddddd=info(0)& "你千辛萬苦尋訪<font color=red>龍珠</font>的下落，卻不料在一條水溝裡被你找到了，趕緊洗洗龍珠挖掘它的魔力吧." 
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
end function 
%>