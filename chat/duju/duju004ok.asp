<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "你不能進行操作，進行此操作必須進入聊天室！"
location.href = "javascript:history.go(-1)"
</script>
<%end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
num=trim(request.form("num"))
for t=1 to len(num)
abc=mid(num,t,1)
if abc<>"0" and abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('操作錯誤，下注請使用數字！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
next
yinliang=abs(int(num))
if Request.Form("submit")="大" then
tt="大"
end if
if Request.Form("submit")="小" then
tt="小"
end if
if Request.Form("submit")="單" then
tt="單"
end if
if Request.Form("submit")="雙" then
tt="雙"
end if
select case tt
	case "大"
		duboimg="<img src='../jhimg/da.gif'>"
	case "小"
		duboimg="<img src='../jhimg/xiao.gif'>"
	case "單"
		duboimg="<img src='../jhimg/dan.gif'>"
	case "雙"
		duboimg="<img src='../jhimg/shuang.gif'>"
case else
	Response.Write "<script language=JavaScript>{alert('押操作為：大、小、單、雙！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end select
if yinliang<10 or yinliang>500 then
		Response.Write "<script language=JavaScript>{alert('最少押：10點戰鬥經驗，最多500點戰鬥經驗！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
'開始判斷
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select allvalue FROM 用戶 WHERE id="&info(9),conn
	if rs("allvalue")<yinliang then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('你有這麼多的存點嗎？！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.close
	tmprs=conn.execute("Select count(*) As 數量 from 點數賭局 where 點數<>0")
	durs=tmprs("數量")
set tmprs=nothing
	if durs>=5 then
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('共["&durs&"]注正在等待開局，稍後再下注！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.open "select top 1 姓名 FROM 點數賭局 WHERE 身份='莊家'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('現在沒有莊家，不能進行點數賭局！！');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.close
	rs.open "select top 1 * FROM 點數賭局 WHERE 姓名='"& info(0) &"'",conn
	if not(rs.eof) or not(rs.bof) then
		if rs("身份")="莊家" then
			temp=info(0)&"你現在是莊家呀，你要搞什麼呀!"
		else
			temp=info(0)&"你壓["&rs("押大小")&"]"&rs("點數")&"等著開吧！"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	conn.execute "update 用戶 set allvalue=allvalue-"&yinliang&" where id="&info(9)
	conn.execute "insert into 點數賭局(姓名,身份,點數,押大小,時間) values ('"&info(0) &"','玩家',"&yinliang&",'"&tt&"',now())"	
	tmprs=conn.execute("Select count(*) As 數量 from 點數賭局 where 押大小='大'")
	dars=tmprs("數量")
	tmprs=conn.execute("Select count(*) As 數量 from 點數賭局 where 押大小='小'")
	xiaors=tmprs("數量")
	tmprs=conn.execute("Select count(*) As 數量 from 點數賭局 where 押大小='單'")
	danrs=tmprs("數量")
	tmprs=conn.execute("Select count(*) As 數量 from 點數賭局 where 押大小='雙'")
	shuangrs=tmprs("數量")
	tmprs=conn.execute("Select count(*) As 數量 from 點數賭局 where 點數<>0")
	durs=tmprs("數量")


'開局了
if durs>=5 then
	Randomize
	m1 = Int(6 * Rnd + 1)
	Randomize
	m2 = Int(6 * Rnd + 1)
	Randomize
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'查找莊家
rs.open "select 姓名 FROM 點數賭局 WHERE 身份='莊家'",conn
zhuangjia=rs("姓名")
rs.close

'豹子處理
if m1=m2 and m2=m3 and m3=m1 then
	rs.open "select 點數 FROM 點數賭局 WHERE 點數<>0 and 身份<>'莊家'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("點數")
	rs.movenext
	loop
	rs.close
	conn.execute "update 用戶 set allvalue=allvalue+"&yingyin&" where 姓名='"& zhuangjia &"'"
	xiazhu="莊家開：<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>計：<font color=red>"& sjdubo &"</font>點!莊家開出豹子  通殺！莊家：<font color=red>"&zhuangjia&"</font>贏得：<font color=red>"&yingyin&"</font>點存點！"
		Application.Lock
	Application("ljjh_dubozhang3")=""
Application.UnLock
set rs=nothing
		conn.close
		set conn=nothing
		call say(xiazhu)
Response.Write "<script Language=Javascript>alert('下注"&yinliang&"點泡點成功！');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if

'初始化數據
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'開單雙處理
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src='../jhimg/shuang.gif'>"
	rs.open "select 點數,姓名 FROM 點數賭局 WHERE 押大小='雙'",conn
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("點數")
		shuangname=shuangname&rs("姓名")&" "
		conn.execute "update 用戶 set allvalue=allvalue+"&rs("點數")*2&" where 姓名='"& rs("姓名") &"'"
		conn.execute "update 用戶 set allvalue=allvalue-"& rs("點數") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from 點數賭局 where  押大小='雙'"
else
	danshuang="<img src='../jhimg/dan.gif'>"
	rs.open "select 點數,姓名 FROM 點數賭局 WHERE 押大小='單'",conn
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("點數")
		danname=danname&rs("姓名")&" "
		conn.execute "update 用戶 set allvalue=allvalue+"&rs("點數")*2&" where 姓名='"& rs("姓名") &"'"
		conn.execute "update 用戶 set allvalue=allvalue-"& rs("點數") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from 點數賭局 where  押大小='單'"
end if


'對開大小處理
if sjdubo<=10 then
	daxiao="<img src='../jhimg/xiao.gif'>"
	rs.open "select 點數,姓名 FROM 點數賭局 WHERE 押大小='小'",conn
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("點數")
		xiaoname=xiaoname&rs("姓名")&" "
		conn.execute "update 用戶 set allvalue=allvalue+"&rs("點數")*2&" where 姓名='"& rs("姓名") &"'"
		conn.execute "update 用戶 set allvalue=allvalue-"& rs("點數") &" where 姓名='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "delete * from 點數賭局 where  押大小='小'"

else
	daxiao="<img src='../jhimg/da.gif'>"
	rs.open "select 點數,姓名 FROM 點數賭局 WHERE 押大小='大'",conn
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("點數")
		daname=daname&rs("姓名")&" "
		conn.execute "update 用戶 set allvalue=allvalue+"&rs("點數")*2&" where 姓名='"& rs("姓名") &"'"
		conn.execute "update 用戶 set allvalue=allvalue-"& rs("點數") &" where 姓名='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from 點數賭局 where  押大小='大'"
end if

'對剩下輸的用戶點數
	rs.open "select 點數 FROM 點數賭局 WHERE 點數<>0 and 身份<>'莊家'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("點數")
	rs.movenext
	loop
	rs.close
	conn.execute "update 用戶 set allvalue=allvalue+"&yingyin&" where 姓名='"& zhuangjia &"'"
	conn.execute "delete * from 點數賭局 where  姓名<>''"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
'shuangyinliang=0
'shuangname=""
'danyinliang=0
'danname=""
'xiaoyinliang=0
'xiaoname=""
'dayinliang=0
'daname=""
	duboname=shuangname&danname&xiaoname&daname

	
	xiazhu="莊家開：<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>計：<font color=red>"& sjdubo &"</font>點!"&danshuang&daxiao&"總下注：<font color=red>"&zong&"</font>點泡點，莊家：<font color=red>["&zhuangjia&"]</font>收入：<font color=red>"&yingyin &"</font>點泡點,賠出：<font color=red>"&pei&"</font>點泡點，合計：<font color=red>"&yingyin-pei&"</font>點泡點,共有：<font color=red>"&duboname&"</font>玩家幸運押中！"
Application.Lock
	Application("ljjh_dubozhang3")=""
Application.UnLock
	set rs=nothing
	conn.close
	set conn=nothing
	call say(xiazhu)
Response.Write "<script Language=Javascript>alert('下注"&yinliang&"點泡點成功！');location.href = 'javascript:history.go(-1)';</script>"
Response.End

end if
xiazhu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>心疼地把自己努力泡來的存點拿出:<font color=red>"& yinliang &"</font>點，我押"& duboimg &"！，一定中的！現在下注如下：押大：<font color=red>"& dars &"</font>個，押小:<font color=red>"& xiaors &"</font>個,押單：<font color=red>"&danrs&"</font>個,押雙:<font color=red>"&shuangrs&"</font>個！還差:<font color=red>"&(5-durs)&"</font>個開局！"

	set rs=nothing	
	conn.close
	set conn=nothing

	call say(xiazhu)
Response.Write "<script Language=Javascript>alert('下注"&yinliang&"點泡點成功！');location.href = 'javascript:history.go(-1)';</script>"
Response.End

sub say(xiazhu)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)=info(0)
sd(195)=info(0)
sd(199)="<font color=green>【經驗賭局】</font><font color=#FFC01F>"& xiazhu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
%>