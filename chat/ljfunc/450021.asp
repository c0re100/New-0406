<%'取法力
function getfali(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 法力點,結算日期,法力,操作時間 from 用戶 where id="&info(9),conn
if DateDiff("s",rs("操作時間"),now())<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('為了防止在法力寶瓶亂刷，存取法力點等10秒鐘操作!');}</script>"
	Response.End 
end if
bankmoney=rs("法力點")
lastdate=date()-rs("結算日期")
money=rs("法力")
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getyin=info(0) & "你在法力寶瓶並沒有存放法力點,官府代為保管法力寶瓶每天利息為:0.0001%,歡迎存取!"
	else
		getyin=info(0) & "在法力寶瓶裡共存有:"& bankmoney &"點法力,在:"& rs("結算日期") &"存入,按0.0001%利,法力寶瓶裡現在有:"& newbankmoney &"點法力!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('你哪裡有那麼多的法力啊？？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getfali=info(0) & "從法力寶瓶裡取出:"& fn1 &"點法力,法力寶瓶現存有法力:"& newbankmoney-fn1 &"點,好好使用,別被別人吸了!"
conn.execute "update 用戶 set 法力=法力+"  & fn1 & ",法力點="& newbankmoney-fn1 &",結算日期=date(),操作時間=now() where id="&info(9)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>