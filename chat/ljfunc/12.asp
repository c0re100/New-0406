<%'吸星大法
function xxdf(to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('吸錯對像啦，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以吸別人內力，不然就會死翹翹的！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名,內力 from 用戶 where 姓名='"&toname&"'",conn
toname=rs("姓名")
nlto=rs("內力")
randomize timer
s=int(rnd*20)+1
nlto=int((nlto/100))*s
r=int(rnd*4)
Select Case r
Case 1
	xxdf=info(0) & "<bgsound src=wav/xixing.wav loop=1>使用一招吸星大法,吸取"& toname &"內力<font color=red>"& s &"%</font>,共計:<font color=red>"& nlto &"</font>點!但不小心全部丟失了!"
Case 2
	xxdf=info(0) & "<bgsound src=wav/xixing.wav loop=1>想吸取"& toname &"內力,沒有成功！自己失去內力<font color=red>"& s &"%</font>,共計:<font color=red>"& nlto &"</font>點!"& toname &"哈,知道我厲害了吧！"
	conn.execute "update 用戶 set 內力=內力-"&nlto&" where id="&info(9)
case 3
rs.close
	rs.open "select 銀兩 from 用戶 where id="&info(9),conn
	if rs("銀兩")>50000 then
		conn.execute "update 用戶 set 銀兩=銀兩-50000 where id="&info(9)
		xxdf=info(0) & "<bgsound src=wav/qt.wav loop=1>偷" & toname & "的內力被發現,隻好上下打點花了<font color=red>5萬</font>塊,逃過此劫,幸運呀!"
	else
		conn.execute "update 用戶 set 狀態='獄' where id="&info(9)
		xxdf=info(0) & "<bgsound src=wav/oh_no.wav loop=1>剛剛把手搭在" & toname & "肩膀上,就被官府的人發現了,立即被捉去坐牢了"
		call boot(info(0))
	end if
case else
	xxdf=info(0) & "<bgsound src=wav/xixing.wav loop=1>使用一招吸星大法,吸取"& toname &"內力<font color=red>"& s &"%</font>,共計:"& nlto &"點!<font color=red>"& toname &"</font>點，氣得"& toname &"吐血！"
	conn.execute "update 用戶 set 內力=內力+"&nlto&" where id="&info(9)
	conn.execute "update 用戶 set 內力=內力-"&nlto&" where 姓名='"&toname&"'"
End Select
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>
