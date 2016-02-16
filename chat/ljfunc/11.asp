<%'偷錢
function touqian(to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('偷錢對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以偷錢，不然就會死翹翹的！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩,體力,內力,grade from 用戶 where id="&info(9),conn
inyin=rs("銀兩")
if rs("體力")<300 or rs("內力")<300  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('【"&info(0)&"】你現在的內力或體力小於300請不要進行偷錢命令操作！');</script>"
	Response.End
end if
rs.close
rs.open "select 姓名,會員等級,銀兩,grade from 用戶 where 姓名='"&toname&"'",conn
toname=rs("姓名")
jhhy=rs("會員等級")
yin=rs("銀兩")
randomize timer
'會員2級以上偷錢成功小於5%！
if jhhy<>0 then
	s=int(rnd*5)+1
else
	s=int(rnd*15)+1
end if
yin=int((yin/100))*s
r=int(rnd*10)

randomize timer
kaa=int(rnd*99)+1
if yin>=1000000 then
randomize timer
yin=1000000
yin=killer+kaa
end if

Select Case r
Case 1
	conn.execute "update 用戶 set 銀兩=銀兩-"&yin&" where 姓名='"&toname&"'"
	touqian=info(0) & "<bgsound src=wav/xiaotou.wav loop=1>發了一招飛龍探雲手,偷取"& toname &"銀兩"& s &"%,共計:"& yin &"兩!但不小心全部丟失了!"
Case 2
	touqian=info(0) & "<bgsound src=wav/xiaotou.wav loop=1>發了一招飛龍探雲手,怎奈"& toname &"與官府認識,"& info(0) &"連忙賠禮道歉,灰溜溜的走了!"
Case 3
	conn.execute "update 用戶 set 體力=體力-1500 where id="&info(9)
	touqian=info(0) & "<bgsound src=wav/oh_no.wav loop=1>想偷取"& toname &"的銀兩,不過被大家發現了,體力下將<font color=red>1500</font>點!<img src='img/daren.gif'>"
Case 4
	if inyin>50000 then
		conn.execute "update 用戶 set 銀兩=銀兩-50000 where id="&info(9)
		touqian=info(0) & "<bgsound src=wav/qt.wav loop=1>偷" & toname & "的錢被發現,隻好上下打點花了<font color=red>5萬</font>塊,逃過此劫,幸運呀!"
	else
		conn.execute "update 用戶 set 狀態='獄',銀兩=0,存款=0 where id="&info(9)
		touqian=info(0) & "<bgsound src=wav/oh_no.wav loop=1>剛剛把手伸進" & toname & "的錢袋,就被官府的人發現了,立即被捉去坐牢了"
		call boot(info(0))
	end if
case else
	conn.execute "update 用戶 set 銀兩=銀兩-"&yin&" where 姓名='"&toname&"'"
	conn.execute "update 用戶 set 銀兩=銀兩+"&yin&" where id="&info(9)
	touqian=info(0) & "<bgsound src=wav/xiaotou.wav loop=1>發了一招飛龍探雲手,偷取"& toname &"銀兩"& s &"%,共計:"& yin &"兩!"& toname &"大叫,我要告你喲!"
End Select
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>
