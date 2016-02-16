<%'同意
function dy(to1)
if trim(Application("ljjh_qqq_sf"))<> trim(info(0)) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]也沒有想娶你啊！！');}</script>"
	Response.End
end if
if trim(Application("ljjh_qqq_td"))<> trim(to1) then
	Response.Write "<script language=JavaScript>{alert('["& to1 &"]也沒有想娶你啊！！');}</script>"
	Response.End
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 門派 FROM 用戶 WHERE id="&info(9),conn
if rs("門派")="官府" then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('為了維護江湖管理、公平競爭,官府禁止結婚！');}</script>"
	Response.End
end if
    conn.execute "update 用戶 set 銀兩=銀兩-5000000,二婚='"& info(0) &"' where 姓名='"& Application("ljjh_qqq_td") &"'"
    conn.execute "update 用戶 set 銀兩=銀兩+5000000,二婚='"& Application("ljjh_qqq_td") &"' where 姓名='"& info(0) &"'"
if info(1)="男" then
dy=Application("ljjh_qqq_td") & "向"& info(0) &"大展媚功,9吻不斷,樂得"&info(0)&"連連叫好，高高興興的收了對方的1000萬禮金，"&Application("ljjh_qqq_td") & "終於讓對方成為自己的小老公，可喜可賀呀！"
else
dy=Application("ljjh_qqq_td") & "向"& info(0) &"說了無數甜言蜜語，並且交了500萬禮金，終於讓對方成為自己的小老婆，可喜可賀呀！小老婆都有了呀真是艷福不淺啊..."
end if

Application("ljjh_qqq_sf")=""
Application("ljjh_qqq_td")=""
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
end function
%>