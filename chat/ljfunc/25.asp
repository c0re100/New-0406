<%'拜師
function bais(to1,toname)
if toname="大家" or toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('拜師對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
if info(3)<2 then
	Response.Write "<script language=JavaScript>{alert('等級太低，要2級以上才可以拜師！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 師傅,銀兩,等級 from 用戶 where id="&info(9),conn
sf=rs("師傅")
if sf=toname then
	Response.Write "<script language=JavaScript>{alert('["& toname&"]已經是你師傅了');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End

end if
if rs("師傅")<>"" and rs("師傅")<>"無" then
	Response.Write "<script language=JavaScript>{alert('想拜["& toname&"]為師，請與現在師傅脫離關繫！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if rs("銀兩")<50000 then
	Response.Write "<script language=JavaScript>{alert('你沒有5萬兩["& toname&"]不想收你為弟子！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
dj=rs("等級")
rs.close
rs.open "select 等級 FROM 用戶 WHERE 姓名='"&toname&"'",conn
if rs("等級")<30 or dj>rs("等級") then
	Response.Write "<script language=JavaScript>{alert('["& toname&"]還不夠30級呀,不能收弟子！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
bais=info(0) & "正準備向"& toname &"交納了5萬塊拜師費，提出了拜師申請，也不知道人家同意不！"
Application("ljjh_bais_sf")=toname
Application("ljjh_bais_td")=info(0)
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
