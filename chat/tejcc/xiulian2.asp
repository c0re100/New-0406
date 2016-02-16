<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
Response.Write "<script language=JavaScript>{alert('滾吧，你想做什麼？想搗亂嗎？！');window.close();}</script>"
Response.End 
end if
un=info(0)
grade=info(3)
myid=info(9)
eec=info(4)
if info(4)=0 then bbc=0 else bbc=1
mg=Request.QueryString("mg")
if instr(mg,"'")<>0 then Response.end
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
connstr=Application("ljjh_usermdb")
conn.open connstr
sql= "select 操作時間,會員等級 from 用戶 where 姓名='"&un&"'"
Set Rs=conn.Execute(sql)
sj=DateDiff("s",rs("操作時間"),now())
if sj<60 and rs("會員等級")=0 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=60-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒再操作！加入會員無此限制！');location.href = 'javascript:history.go(-1)';}</script>}</script>"
	Response.End
end if
sql="select 等級,修練法力,魔法 from 使用技能 where 姓名='"&un&"' and 技能='"&mg&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
sql= "select * from 職業技能 where 技能='"&mg&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
learn=""
else
energy=rs("修練法力")
proviso=rs("修練條件")
basemp=rs("消耗法力")
baseat=rs("基本攻擊")
especial=rs("魔法")
eepic=rs("圖檔")
atdeclaration=rs("施法說明")
sql= "select id,操作時間 from 用戶 where 姓名='"&un&"' and 法力>="&energy&" and "&proviso
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
provisotxt=replace(proviso,"and","並且")
provisotxt=replace(provisotxt,"or","或者")
provisotxt=replace(provisotxt,">=","不低於")
provisotxt=replace(provisotxt,"<=","不高於")
provisotxt=replace(provisotxt,">","大於")
provisotxt=replace(provisotxt,"<","小於")
provisotxt=replace(provisotxt,"=","為")
learn="修練時失敗，需法力"&energy&"，需要條件為："&provisotxt
else
conn.Execute "insert into 使用技能(姓名,技能,等級,修練法力,消耗法力,基本攻擊,魔法,施法說明,圖檔,aszcc) values('"&un&"','"&mg&"',1,"&energy&","&basemp&","&baseat&",'"&especial&"','"&atdeclaration&"','"&eepic&"','"&bbc&"')"
conn.execute "update 用戶 set 法力=法力-"&energy&",操作時間=now() where 姓名='"&un&"'"
learn="修練"&mg&"第1級成功。耗用法力"&energy
end if 
end if
elseif rs("等級")<200 then
energy=rs("修練法力")
agrade=rs("等級")
sql="select id from 用戶 where 姓名='"&un&"' and 法力>="&energy*agrade*agrade
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
learn="修練時失敗，需法力"&energy*agrade*agrade
else
conn.Execute "update 用戶 set 法力=法力-"&energy*agrade*agrade&" where 姓名='"&un&"'"
conn.Execute "update 使用技能 set 等級=等級+1,基本攻擊=基本攻擊*1 where 姓名='"&un&"' and 技能='"&mg&"'"
learn="修練"&mg&"第"&agrade&"級成功。耗用法力"&energy*agrade*agrade
end if
else learn="修練"&mg&"至最高了無法再行修練。"
end if
set rs=nothing
conn.Close
set conn=nothing
%>
<html>
<head>
<title>修習技能</title>
</head>
<meta http-equiv="cnntent-Type" cnntent=text/html; charset=big5>
<body bgcolor='#996699' background="../bk.jpg">
<div align=center> <font color="#FFFFFF" size="2">修練技能</font> 
<hr>
<font color="#ffffff" size="2"><%=learn%></font> <br>
<input type=button onClick="javascript:location.href='xiulian.asp'" value='返回' name="button">
</div>
</html>