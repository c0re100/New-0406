<%'跟蹤私聊
function tracksl()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 姓名,天眼 FROM 用戶 WHERE id="&info(9),conn
if rs("天眼")<>1 then 
tracksl="[系統]對" & info(0) & "說：沒這個命令啊！你是不是搞錯了？"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
else
 if session("slshow")=0 then 
  session("slshow")=1 
Response.Write "<script language=JavaScript>{alert('現在你的天眼已開通，你可以看到對方的所有話！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
'tracksl="<font color=red><b>現在你的天眼已開通，你可以看到對方的所有話!</b></font>" 
  else 
Response.Write "<script language=JavaScript>{alert('你的天眼已開通了呀！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
'tracksl="<font color=red><b>你的天眼已開通了呀!</b></font>" 
  end if 
end if 
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>
