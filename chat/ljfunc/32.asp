<%'取錢
function getyin(fn1)
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 存款,結算日期,銀兩,操作時間 from 用戶 where id="&info(9),conn
if DateDiff("s",rs("操作時間"),now())<10 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('為了防止在銀行刷錢，存取錢等10秒鐘操作!');}</script>"
	Response.End 
end if
bankmoney=rs("存款")
lastdate=date()-rs("結算日期")
money=rs("銀兩")
moneyok=int(money)+abs(fn1)
if moneyok>2000000000 then
	Response.Write "<script language=JavaScript>{alert('你的現金太多就快超過了20億了，為防止丟錢現像請您把錢存些進股市吧！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
newbankmoney=int(bankmoney+bankmoney*0.0001*lastdate)
fn1=abs(fn1)
if fn1<=0 then
	if bankmoney=<0 then
		getyin=info(0) & "你在江湖錢莊並沒有存款,錢莊每天利息為:0.0001%,歡迎存取!"
	else
		getyin=info(0) & "在錢莊存有:"& bankmoney &"兩,在:"& rs("結算日期") &"存入,按0.0001%利,銀行現在有:"& newbankmoney &"兩!"
	end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	exit function
end if
if newbankmoney<fn1 then
	Response.Write "<script language=JavaScript>{alert('你哪裡有那麼多的錢？？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
getyin=info(0) & "向錢莊取錢:"& fn1 &"兩,錢莊現有銀兩:"& newbankmoney-fn1 &"兩,拿好你的錢,別被搶了!"
conn.execute "update 用戶 set 銀兩=銀兩+"  & fn1 & ",存款="& newbankmoney-fn1 &",結算日期=date(),操作時間=now() where id="&info(9)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>