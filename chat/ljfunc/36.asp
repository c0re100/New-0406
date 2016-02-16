<%'暴威力豆baodou
function baodou()
Set rs=Server.CreateObject("ADODB.RecordSet")
Set Conn=Server.CreateObject("ADODB.Connection")
conn.open Application("ljjh_usermdb")
rs.open "select 身份,grade,泡豆點數,暴豆時間 FROM 用戶 WHERE id="&info(9),conn
if trim(rs("身份"))="護法" and rs("grade")=3 then
	doudi=500
else
	doudi=1000
end if
if rs("泡豆點數")<doudi then
	Response.Write "<script language=JavaScript>{alert('你哪裡有威力豆？1個豆可以在一小內威力漲3倍！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
sj=DateDiff("n",rs("暴豆時間"),now())
if sj<=60 then
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('"& info(0) & "威力時間未過，還剩"& ss &"分鐘!');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
conn.execute "update 用戶 set 暴豆時間=now(),泡豆點數=泡豆點數-"& doudi &" where id="&info(9)
baodou=info(0) & "祝賀你暴威力豆操作成功，從現在開始你的攻擊力會比原來大3倍！"
rs.close
set rs=nothing	
conn.close
set conn=nothing
end function
%>