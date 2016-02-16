<%'廣招弟子
function jiamen(fn1,to1,toname)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 身份 FROM 用戶 WHERE id="&info(9),conn
if rs("身份")<>"掌門" and rs("身份")<>"副掌門" and rs("身份")<>"長老" and  rs("身份")<>"護法" and rs("身份")<>"堂主" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你是什麼身份，不夠資格!');}</script>"
	Response.End
end if
rs.close	
rs.open "select 幫派基金,適合 FROM 門派 WHERE 門派='" & info(5) & "'",conn
if rs("幫派基金")<15000000 then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('你門派的基金太少了，至少得有1500萬啊！');}</script>"
	Response.End
end if
shihe=trim(rs("適合"))
rs.close
set rs=nothing
conn.close
set conn=nothing
Application("menpai")=info(5)
jiamen="["&info(5)&"]["&info(0)&"]向大家叫到：["&info(5)&"]廣招弟子,人多力量大啊,各位大俠別猶豫了,快快加入本派,本派招收<font color=#009900>"&shihe&"弟子</font>,先加入"&Application("menpai")&"的有1000萬銀兩啊  <input type=button style='FONT-SIZE: 9pt'  value='加入' onClick=javascript:;disabled=1;window.open('jiamen.asp','d')>"
end function
%>
