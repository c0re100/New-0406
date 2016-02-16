<%'閉目養神
function bimu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 內力,體力,等級,體力加,操作時間 FROM 用戶 WHERE id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');}</script>"
	Response.End
end if
if rs("內力")<110  then
	Response.Write "<script language=JavaScript>{alert('需要體力100才可以操作！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("體力")<rs("等級")*240+5260+rs("體力加") then
	conn.execute "update 用戶 set 體力=體力+1500,內力=內力-100,操作時間=now() where id="&info(9)
	bimu=info(0) & "<bgsound src=wav/dz.wav loop=1>進行練功打坐,內力失去<font color=red>-100</font>點,體力提升<font color=red>+1500</font>點,有得必有失嘛!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("體力加")<=rs("等級")*1600 then
	conn.execute "update 用戶 set 體力加=體力加+9,內力=內力-100,操作時間=now() where id="&info(9)
	bimu=info(0) & "<bgsound src=wav/dz.wav loop=1>進行練功打坐,內力失去-100體力上限提升+9,有得必有失嘛!"
else
	Response.Write "<script language=JavaScript>{alert('現在你的上限滿了，等升了級再練吧');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
