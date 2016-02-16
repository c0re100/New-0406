<%
'官府獎勵
function jianglila(fn1,to1,toname)
if info(2)<10 then
	Response.Write "<script language=JavaScript>{alert('想作什麼呀，你的管理等級可不夠呀！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 銀兩 FROM 用戶 WHERE id="&info(9),conn
fn1=int(abs(fn1))
if fn1>500000 or fn1<1000  then
Response.Write "<script language=JavaScript>{alert('官府獎勵不可以超過50萬少不可少於1000的！');}</script>"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.End
end if
conn.execute "update 用戶 set 銀兩=銀兩+" & fn1 & " where 姓名='" & toname & "'"
jianglila="因"& toname &"對江湖有貢獻，"&info(0) & "把官府設立的獎金發給 " & toname &" "& fn1 & "兩作為獎勵！"
rs.close
conn.close
set rs=nothing	
set conn=nothing

end function
%>