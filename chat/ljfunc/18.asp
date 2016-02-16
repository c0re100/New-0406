<%'查看狀態
function getstat(to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select * FROM 用戶 WHERE 姓名='"&toname&"'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('你所查看的人不存在！！！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if toname=info(0) or info(2)>=9  then
	getstat=toname & "," & rs("性別") & ",伴侶:" & rs("配偶") & ",門派:" &rs("門派") & ",身份:" &rs("身份") &  ",武功:" &rs("武功") & ",內力:" & rs("內力") & ",體力:" & rs("體力") & ",攻擊力:" & rs("攻擊") & ",防御力:" & rs("防御")&",當前ip:" & rs("ip") & ",注冊ip:" & rs("regip") & ",最後ip:" & rs("lastip") & ",狀態:" & rs("狀態") & ",信箱:" & rs("信箱") & ",魅力:" & rs("魅力") & ",帖子數:" & rs("帖子數") & ",道德:" & rs("道德") & ",管理等級:" & rs("grade") & ",累積分:" & rs("allvalue") & ",月積分:" & rs("mvalue") & ",登錄:" & rs("登錄") & ",銀兩:" & rs("銀兩")
else
	if  info(2)>rs("grade") then
		Response.Write "<script language=JavaScript>{alert('對方管理級別比你高，嚴禁查看！！！');}</script>"
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.End
	end if
	getstat=toname & "," & rs("性別") & ",伴侶:" & rs("配偶") & ",門派:" &rs("門派") & ",身份:" &rs("身份")
end if
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
