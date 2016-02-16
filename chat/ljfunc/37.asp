<%'啞穴
function ya(to1,toname,fn1)
if info(2)<4 then
	Response.Write "<script language=JavaScript>{alert('想作什麼呀，你的管理等級可不夠呀！');}</script>"
	Response.End
end if
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('點啞穴對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(2)<3 then
	Response.Write "<script language=JavaScript>{alert('你是什麼身份？一不是官府，二不是護法級以上高手！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select grade,門派 FROM 用戶 WHERE id="&info(9),conn
grade=rs("grade")
menpai=rs("門派")
denji=rs("grade")
rs.close
rs.open "select 門派,grade from 用戶 where 姓名='"&toname&"'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('沒有這個人？你是不是看錯了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if denji<6 and menpai<>rs("門派") then
	Response.Write "<script language=JavaScript>{alert('搞錯了["&toname&"]也不是你們門派的呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("grade")<grade  then
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
application("ljjh_dianxuename")=application("ljjh_dianxuename")&toname&"|"&sj2&"|"&";"
	ya=info(0)& "對" & toname & "使用了啞穴術，" & toname & "獃獃地不動了  [" & fn1 & "]"
else
	ya=info(0) & "對" & toname & "使用了啞穴術，可是你的等級不如人家？沒辦法！"
end if
'記錄
conn.execute "insert into 操作記錄(時間,姓名1,姓名2,方式,數據) values (now(),'"& info(0) &"','"& toname &"','啞穴','"& fn1 & "')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>