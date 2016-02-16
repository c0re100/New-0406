<%'查ip
function seeip(to1,toname)
if info(2)<7 then
	Response.Write "<script language=JavaScript>{alert('管理需要7級的才可以查看ip的！');}</script>"
	Response.End
end if
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('查ip對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'查用戶ip
rs.open "select lastip,regip FROM 用戶 WHERE  姓名='"&toname&"'",conn
ip1=rs("lastip")   '最後ip
ip2=rs("regip")    '注冊ip
sip=split(rs("lastip"),".")
sip1=split(rs("regip"),".")
num=cint(sip(0))*256*256*256+cint(sip(1))*256*256+cint(sip(2))*256+cint(sip(3))-1
'查最後ip的地區
rs.close
rs.open "select 國家,城市 FROM 地址 WHERE ip1<="& num &" and ip2>="&num,conn
if rs.eof or rs.bof then
	country="亞洲"
	city="未知"
else
	country=rs("國家")
	city=rs("城市")
end if
if country="" then country="中國"
if city="" then city="未知"
last="地區:"& country &" 城市:"& city
rs.close
'查注冊ip的地區
num=cint(sip1(0))*256*256*256+cint(sip1(1))*256*256+cint(sip1(2))*256+cint(sip1(3))-1
rs.open "select 國家,城市 FROM 地址 WHERE ip1<="& num &" and ip2>="&num,conn
if rs.eof or rs.bof then
	country="亞洲"
	city="未知"
else
	country=rs("國家")
	city=rs("城市")
end if
if country="" then country="中國"
if city="" then city="未知"
reg="地區:"& country &" 城市:"& city
seeip="[查ip]"& toname &"的現在ip地址為:"&ip1&",鋻定為："&last&"  注冊ip地址為:"&ip2 &"鋻定為："&reg
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>