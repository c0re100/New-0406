<%'練武
function lianwu()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 內力,體力,武功,等級,武功加,操作時間 FROM 用戶 WHERE id="&info(9),conn
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
if rs("體力")<210 or rs("內力")<110 then
	Response.Write "<script language=JavaScript>{alert('需要體力200，內力100才可以');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("武功")<rs("等級")*1280+3800+rs("武功加") then
	conn.execute "update 用戶 set 武功=武功+130,體力=體力-200,內力=內力-100,操作時間=now() where id="&info(9)
	lianwu=info(0) & "<bgsound src=wav/dz.wav loop=1>進行練功習武,體力失去-200,內力失去-100,武功提升+130,有得必有失嘛!"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if rs("武功加")<=rs("等級")*1000 then
	conn.execute "update 用戶 set 武功加=武功加+9,體力=體力-200,內力=內力-100,操作時間=now() where id="&info(9)
	lianwu=info(0) & "<bgsound src=wav/dz.wav loop=1>進行練功習武,體力失去-200,內力失去-100,武功上限提升+9,有得必有失嘛!"
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