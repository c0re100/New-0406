<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/liugrade")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if
id=lcase(trim(request("id")))
if InStr(id,"oR")<>0 or InStr(id,"Or")<>0 or inStr(id,"&")<>0 or inStr(id,"&&")<>0 or InStr(id,"OR")<>0 or InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select id from 物品 where 物品名='龍珠' and 數量>0 and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('改良武器需要龍珠！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
idd=rs("id")
conn.execute "update 物品 set 數量=數量-1 where id="&idd
rs.close
rs.open "select 類型,等級,攻擊,防御,物品名 from 物品 where id="&id&" and 數量>0 and 裝備=false and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器正裝備著呢，若要升級請解除裝備！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
end if
dj=rs("等級")
gj=rs("攻擊")
fy=rs("防御")
wuname=mid(rs("物品名"),1,2)
randomize timer
r=int(rnd*dj*2)
if r<2 then
if rs("類型")="右手" then
select case wuname
case "良品"
ganliang=Replace(rs("物品名"),"良品","上品")
conn.execute "update 物品 set 物品名='"&ganliang&"',攻擊=攻擊+2000 where id="&id
case "上品"
ganliang=Replace(rs("物品名"),"上品","精品")
conn.execute "update 物品 set 物品名='"&ganliang&"',攻擊=攻擊+3000 where id="&id
case "精品"
ganliang=Replace(rs("物品名"),"精品","極品")
conn.execute "update 物品 set 物品名='"&ganliang&"',攻擊=攻擊+4000 where id="&id
case "極品"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器品質已達到極限不可再改良了，這個龍珠我就幫你保管啦,嘿嘿！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update 物品 set 物品名='良品'&物品名,防御=防御+1000 where id="&id
end select
end if
if rs("類型")="左手"  then
select case wuname
case "良品"
ganliang=Replace(rs("物品名"),"良品","上品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+2000 where id="&id
case "上品"
ganliang=Replace(rs("物品名"),"上品","精品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+3000 where id="&id
case "精品"
ganliang=Replace(rs("物品名"),"精品","極品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+4000 where id="&id
case "極品"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器品質已達到極限不可再改良了，這個龍珠我就幫你保管啦,嘿嘿！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update 物品 set 物品名='良品'&物品名,防御=防御+1000 where id="&id
end select
end if
if  rs("類型")="盔甲" then
select case wuname
case "良品"
ganliang=Replace(rs("物品名"),"良品","上品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+1000 where id="&id
case "上品"
ganliang=Replace(rs("物品名"),"上品","精品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+2000 where id="&id
case "精品"
ganliang=Replace(rs("物品名"),"精品","極品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+3000 where id="&id
case "極品"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器品質已達到極限不可再改良了，這個龍珠我就幫你保管啦,嘿嘿！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update 物品 set 物品名='良品'&物品名,防御=防御+500 where id="&id
end select
end if
if  rs("類型")="頭盔" then
select case wuname
case "良品"
ganliang=Replace(rs("物品名"),"良品","上品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+600 where id="&id
case "上品"
ganliang=Replace(rs("物品名"),"上品","精品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+800 where id="&id
case "精品"
ganliang=Replace(rs("物品名"),"精品","極品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+900 where id="&id
case "極品"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器品質已達到極限不可再改良了，這個龍珠我就幫你保管啦,嘿嘿！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update 物品 set 物品名='良品'&物品名,防御=防御+600 where id="&id
end select
end if
if rs("類型")="雙腳" then
select case wuname
case "良品"
ganliang=Replace(rs("物品名"),"良品","上品")
conn.execute "update 物品 set 物品名='"&ganliang&"',攻擊=攻擊+600,防御=防御+600 where id="&id
case "上品"
ganliang=Replace(rs("物品名"),"上品","精品")
conn.execute "update 物品 set 物品名='"&ganliang&"',攻擊=攻擊+700,防御=防御+700 where id="&id
case "精品"
ganliang=Replace(rs("物品名"),"精品","極品")
conn.execute "update 物品 set 物品名='"&ganliang&"',攻擊=攻擊+800,防御=防御+800 where id="&id
case "極品"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器品質已達到極限不可再改良了，這個龍珠我就幫你保管啦,嘿嘿！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update 物品 set 物品名='良品'&物品名,攻擊=攻擊+500,防御=防御+500 where id="&id
end select
end if
if rs("類型")="裝飾" then
select case wuname
case "良品"
ganliang=Replace(rs("物品名"),"良品","上品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+500 where id="&id
case "上品"
ganliang=Replace(rs("物品名"),"上品","精品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+600 where id="&id
case "精品"
ganliang=Replace(rs("物品名"),"精品","極品")
conn.execute "update 物品 set 物品名='"&ganliang&"',防御=防御+700 where id="&id
case "極品"
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('你的武器品質已達到極限不可再改良了，這個龍珠我就幫你保管啦,嘿嘿！');location.href = 'javascript:history.go(-1)';}</script>"
Response.End
case else
conn.execute "update 物品 set 物品名='良品'&物品名,防御=防御+400 where id="&id
end select
end if
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('改良武器成功');location.href = 'javascript:history.go(-1)';}</script>"
else
rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Write "<script language=JavaScript>{alert('改良武器失敗');location.href = 'javascript:history.go(-1)';}</script>"
end if%>
