<%'解除同盟
function jiecmen(fn1,to1,toname)
if toname=info(0) then
	Response.Write "<script language=JavaScript>{alert('解除同盟對像有錯，請看仔細了！！');}</script>"
	Response.End
exit function
end if
fn1=trim(fn1)
if InStr(fn1,"'")<>0 or InStr(fn1,"`")<>0 or InStr(fn1,"=")<>0 or InStr(fn1,"-")<>0 or InStr(fn1,",")<>0 or InStr(fn1,"\")<>0 or InStr(fn1,"/")<>0 or InStr(fn1,"|")<>0 then 
	Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');}</script>"
	Response.End 
end if
if info(6)<>"掌門" then
	Response.Write "<script language=JavaScript>{alert('解除同盟只有掌門才可以操作！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 同盟 FROM 用戶 WHERE id="&info(9),conn
tongmen1=rs("同盟")
rs.close
rs.open "select 同盟 FROM 用戶 WHERE 門派='"&fn1&"'",conn
if rs.bof and rs.eof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('該門派不存在！');}</script>"
	Response.End

end if
tongmen2=rs("同盟")
if  Instr(1,tongmen1,"|"& fn1 & "|")<>0  then
fn11=Replace(tongmen1,"|"& fn1 &"|","")
fn12=Replace(tongmen2,"|"& info(5) &"|","")
	conn.execute "update 用戶 set 同盟='"& fn11 &"' where 門派='"&info(5)&"'"
	conn.execute "update 用戶 set 同盟='"& fn12 &"' where 門派='"&fn1&"'"
else
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('該門派並未與您門派結過盟！');}</script>"
	Response.End

end if

rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
jiecmen="["&info(5)&"]與["&fn1&"]掌門意見不和正式解除同盟！"
end function
%>
