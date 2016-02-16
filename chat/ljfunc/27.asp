<%'打坐
function dazhuo()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 體力,等級,內力加,內力,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
if rs("體力")<210  then
	Response.Write "<script language=JavaScript>{alert('需200以上的體力才可以！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("內力")<rs("等級")*64+2000+rs("內力加") then
	conn.execute "update 用戶 set 內力=內力+30,體力=體力-200,操作時間=now() where id="&info(9)
	dazhuo=info(0) & "<bgsound src=wav/dz.wav loop=1>進行練功打坐,體力失去-200內力提升+30,有得必有失嘛!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("內力加")<=rs("等級")*800 then
	conn.execute "update 用戶 set 內力加=內力加+9,體力=體力-200,操作時間=now() where id="&info(9)
	dazhuo=info(0) & "<bgsound src=wav/dz.wav loop=1>進行練功打坐,體力失去<font color=red>-200</font>點,內力上限提升<font color=red>+9</font>點,有得必有失嘛!"
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