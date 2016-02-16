<%'投擲
function touzi(fn1,to1,toname)
if toname="" or toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('投擲對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
'判斷殺人數
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 殺人數,保護,門派,等級,體力,內力 from 用戶 where id="&info(9),conn
if rs("殺人數")>=30 then
	
	Response.Write "<script language=JavaScript>{alert('你作惡太多，超過江湖殺人上限30不能再投擲了！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

if rs("體力")<300 or rs("內力")<300  then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('【"&info(0)&"】你現在的內力或體力小於300請不要進行投擲命令操作！');</script>"
	Response.End
end if
if rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('請取消練功保護再操作！');</script>"
	Response.End
end if
if rs("等級")<2 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以吸別人內力，不然就會死翹翹的！');}</script>"
	Response.End
end if
if rs("門派")="出家" then
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Response.Write "<script Language=Javascript>alert('你已經出家了，不可以再殺人了！');location.href = 'javascript:history.go(-1)';</script>"
	Response.End
end if
rs.close
rs.open "select grade,等級,保護,體力 from 用戶 where 姓名='"&toname&"'",conn
if rs("等級")<10 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=javascript>alert('【"&toname&"】是新手要注意保護，不能投擲東西！');</script>"
Response.End
	exit function
end if
if rs("保護")=true then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('對方正在閉關！');</script>"
	Response.End
end if
if rs("體力")<0 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('對方已經死了！');</script>"
	Response.End
end if
rs.close
'開始物品
rs.open "select id,內力,體力,數量 from 物品 where 類型='暗器' and 數量>0 and 物品名='" & fn1 & "' and 擁有者="&info(9) ,conn
if rs.eof or rs.bof  then
	Response.Write "<script language=JavaScript>{alert('你哪裡有["& fn1 &"]這樣物品呀！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
id=rs("ID")
neili=abs(rs("內力"))
tili=abs(rs("體力"))
'物品-1並查為0時刪除！
conn.execute "update 物品 set 數量=數量-1 where id="&id
'開始比較
conn.execute "update 用戶 set 道德=道德-20,體力=體力-" & int(neili/4) & " where id="&info(9)
conn.execute "update 用戶 set 內力=內力-" & neili & ",體力=體力-" & tili & " where 姓名='"&toname&"'"
rs.close
rs.open "select 姓名,體力 from 用戶 where 姓名='"&toname&"'",conn
toname=rs("姓名")
if rs("體力")>-100 then
	touzi=info(0) & "<bgsound src=wav/anqi.wav loop=1>向" & toname & "下" & fn1 & ",使" & toname & "的內力減" & neili & "體力減" & tili & "!自己卻消耗內力:-"& int(neili/4 ) &"點!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
conn.execute "update 用戶 set 殺人數=殺人數+1 where id="&info(9)
touzi=info(0) & "<bgsound src=wav/daipu.wav loop=1>向" & toname & "投擲了" & fn1 & ",使" & toname & "的內力:-" & neili & "體力:-" & tili & "," & toname & "咽氣前說道“查出兇手,為我報仇  ”"
conn.execute "update 用戶 set 狀態='死' where 姓名='"&toname&"'"
'記錄
conn.execute "insert into 人命(死者,時間,兇手,死因) values ('" & toname & "',now(),'" & info(0) & "','"&fn1&"')"
rs.close
set rs=nothing	
conn.close
set conn=nothing
call boot(toname)
end function
%>