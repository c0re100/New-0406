<%'冊封
function cefen(fn1,to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('冊封對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(2)<2 then
	Response.Write "<script language=JavaScript>{alert('管理級太低，要2級以上才可以冊封對方！');}</script>"
	Response.End
end if
fn1=trim(fn1)
if len(fn1)>10 then
	Response.Write "<script language=JavaScript>{alert('冊封身份長度不可大於10個字符！');}</script>"
	Response.End
end if
if instr(fn1,"掌門")<>0 then
	Response.Write "<script language=JavaScript>{alert('他作掌門你作什麼？');}</script>"
	Response.End
end if
if len(fn1)>2 and (instr(fn1,"長老")<>0 or instr(fn1,"護法")<>0 or instr(fn1,"堂主")<>0 or instr(fn1,"副手")<>0 ) then
	Response.Write "<script language=JavaScript>{alert('關於:長老、護法、堂主為系統保留，請不要使用在名號中！');}</script>"
	Response.End
end if
cefeng1=instr(says,"=")
cefeng2=instr(says,",")
cefeng3=instr(says,"grade")
cefeng4=instr(says,"身份")
cefeng5=instr(says,"門派")
cefeng6=instr(says,"官府")
cefeng6=instr(says,"江湖總站")
cefeng6=instr(says,"站長")
cefeng6=instr(says,"副站長")
cefeng7=instr(says,"`")
cefeng8=instr(says,"\")
cefeng9=instr(says,"<")
cefeng10=instr(says,">")
if cefeng1<>0 or cefeng2<>0 or cefeng3<>0 or cefeng4<>0 or cefeng5<>0 or cefeng6<>0 or cefeng7<>0 or cefeng8<>0  or cefeng9<>0 or cefeng10<>0 then
	Response.Write "<script language=JavaScript>{alert('操作錯誤！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派 from 用戶 where id="& info(9),conn
aaa=rs("門派")
rs.close
rs.open "select 姓名,grade,等級,會員等級 from 用戶 where 姓名='"&toname&"' and 門派='" & aaa& "'",conn
toname=rs("姓名")
grade=rs("grade")
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('["& toname &"]也不是你派的弟子你想作什麼？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if info(2)<grade then
	Response.Write "<script language=JavaScript>{alert('["& toname &"]的身份比你高啊！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
select case fn1
case "副手"
if info(2)<5 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的管理級不夠操作！');}</script>"
	Response.End
end if
	if rs("等級")<25 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]還不夠25級，不能封為副掌門!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	tmprs=conn.execute("Select count(*) As 數量 from 用戶 where grade=3 and 身份='副掌門' and 門派='"&info(5)&"'")
	musers=tmprs("數量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=2 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]現在你派的副掌門有2個了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用戶 set 身份='副掌門',grade=3 where 姓名='"&toname&"'"
fn1="副掌門"
case "長老"
if info(2)<4 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的管理級不夠操作！');}</script>"
	Response.End
end if
	if rs("等級")<15 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]還不夠15級，不能封長老!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if

	tmprs=conn.execute("Select count(*) As 數量 from 用戶 where grade=2 and 身份='長老' and 門派='"&info(5)&"'")
	musers=tmprs("數量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=4 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]現在你派的長老有4個了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用戶 set 身份='長老',grade=2 where 姓名='"&toname&"'"

case "護法"
if info(2)<3 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的管理級不夠操作！');}</script>"
	Response.End
end if
	if rs("等級")<10 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]還不夠10級，不能封護法!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if

	tmprs=conn.execute("Select count(*) As 數量 from 用戶 where grade=2 and 身份='護法' and 門派='"&info(5)&"'")
	musers=tmprs("數量")
	set tmprs=nothing
	if isnull(musers) then musers=0
	if musers>=8 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]現在你派的護法有8個了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用戶 set 身份='護法',grade=2 where 姓名='"&toname&"'"

case "堂主"
if info(2)<2 then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你的管理級不夠操作！');}</script>"
	Response.End
end if
	if rs("等級")<5 then
		Response.Write "<script language=JavaScript>{alert('["&toname&"]還不夠5級，不能封堂主!');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
tmprs=conn.execute("Select count(*) As 數量 from 用戶 where grade=2 and 身份='堂主' and 門派='"&info(5)&"'")
musers=tmprs("數量")
set tmprs=nothing
if isnull(musers) then musers=0
	if musers>=12 then
		Response.Write "<script language=JavaScript>{alert('["&info(0)&"]現在你派的堂主有12個了，不要再封了！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
conn.execute "update 用戶 set 身份='堂主',grade=2 where 姓名='"&toname&"'"

case else
if aaa="官府" and info(2)>=10 then
conn.execute "update 用戶 set 身份='"& fn1 &"' where 姓名='"&toname&"'"
else
	conn.execute "update 用戶 set 身份='"& fn1 &"',grade=1 where 姓名='"&toname&"'"
end if
end select
cefen=aaa&"派掌門："& info(0) & "冊封" & toname & "為" & aaa & "的<font color=red><b>" & fn1 &"</b></font>成功,大家祝賀！"
'記錄
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','冊封','"& cefen & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>