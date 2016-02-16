<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="select 師傅,武功加,體力,知質,等級,操作時間 from 用戶 where ID=" & info(9)
set rs=conn.execute(sql)
sj=DateDiff("n",rs("操作時間"),now())
if sj<3 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=3-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]分,再操作！');location.href = 'shifu.asp';}</script>"
	Response.End
end if
wushu=rs("武功加")
tili=rs("體力")
zhizi=rs("知質")
sf=rs("師傅")
if sf="無"  then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你還沒有師傅啊，去拜師吧！');location.href = 'shifu.asp';}</script>"
Response.End
end if
if tili<110 or zhizi<20 then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
%>
<script language="vbscript">
  MsgBox "你的內力小於100或者資質小於20，就想練上乘的武功，先把基本功練好吧"
  location.href = "shifu.asp"
</script><%
else
if rs("武功加")<=rs("等級")*1000 then
	conn.execute "update 用戶 set 體力=體力-100,知質=知質-20,武功加=武功加+10,操作時間=now() where id="&info(9)	
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('現在你的上限滿了，等升了級再練吧');location.href = 'shifu.asp';}</script>"
	Response.End
end if
'conn.execute "update 用戶 set 體力=體力-100,知質=知質-20,武功加=武功加+10 where id="&info(9)
%>
<script language="vbscript">
  MsgBox "經過師傅的調教，你花掉資質20點，你的練武極限上升10點。"
  location.href = "shifu.asp"
</script><%
end if
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing%>

