<%'加入門派
function join(fn1)
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 or InStr(fn1,"\")<>0 or InStr(fn1,"/")<>0 or InStr(fn1,"|")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派,性別 FROM 用戶 WHERE id="&info(9),conn
if rs("門派")<>"無" and rs("門派")<>"未定" and rs("門派")<>"遊俠" then
	Response.Write "<script language=JavaScript>{alert('請離開你現在的門派再說吧？');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sex=rs("性別")
rs.close
rs.open "select 適合 FROM 門派 WHERE 門派='" & fn1 & "' and 門派<>'官府'",conn
if rs.eof or rs.bof then
	Response.Write "<script language=JavaScript>{alert('江湖中沒有["& fn1 &"]這樣的門派.');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if trim(sex)=trim(rs("適合")) or trim(rs("適合"))="所有" then
	conn.execute "update 用戶 set 門派='" & fn1 & "', 身份='弟子',grade=1 where id="&info(9)
	conn.execute "update 門派 set 弟子數=弟子數+1 where 門派='" & fn1 & "'"
	join="<font color=ffcc00> " & info(0) & " </font>你成功地加入了<font color=red>" & fn1 & "</font>這個門派!"
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	exit function
end if
Response.Write "<script language=JavaScript>{alert('["&fn1&"]這個門派不適合你呀！');}</script>"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>