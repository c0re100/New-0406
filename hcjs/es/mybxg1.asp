<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(4)=0 then 
aaao=0
else
aaao=1
end if

if Session("ljjh_inthechat")<>"1" then 
	Response.Write "<script language=JavaScript>{alert('你不能進行操作，進行此操作必須進入聊天室！');window.close();}</script>"
	Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/es")=0 then 
Response.Write "<script language=javascript>{alert('對不起，程序拒絕您的操作！！！\n     按確定返回！');parent.history.go(-1);}</script>" 
Response.End 
end if

id=lcase(trim(request.querystring("id")))
sl=abs(int(Request.form("wpsl")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
myname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 對像,物品名,類型,內力,體力,防御,攻擊,銀兩,二手價,擁有者,說明,sm,等級,堅固度 from 交易市場 where id=" & id & " and 數量>="&sl,conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：你家保險櫃裡有這麼多東西嗎？！！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
duixiang=rs("對像")
wpname=rs("物品名")
lx=rs("類型")
nl=rs("內力")
tl=rs("體力")
fy=rs("防御")
gj=rs("攻擊")
say=rs("說明")
sm=rs("sm")
dj=rs("等級")
jgd=rs("堅固度")
yin=rs("銀兩")
'sl=rs("數量")
esmoney=rs("二手價")
toname=rs("擁有者")
if toname<>info(9) then
	conn.close
	set rs=nothing
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：呵。。笑話，想作什麼呀，想偷人家東西！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
rs.close
rs.open "select 銀兩,操作時間 from 用戶 where id="&info(9),conn
sj=DateDiff("s",rs("操作時間"),now())
if sj<10 then
	s=10-sj
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('對不起系統限制，請等["&s&"秒鐘]再操作！');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End
end if	
if rs("銀兩") < int((sl*yin)*0.05) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：搞錯了，自己沒有帶夠錢東西取不走！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
conn.Execute "update 用戶 set 銀兩=銀兩-"& int((sl*yin)*0.05) &",操作時間=now() where id="&info(9)
'conn.execute "delete * from 交易市場 where  id="&id
conn.execute "update 交易市場 set 數量=數量-"& sl &" where id="&id
rs.close
rs.open "select 數量 from 交易市場 where id=" & id ,conn
if rs("數量")=0 then
conn.execute "delete * from 交易市場 where  id="&id
end if
rs.close
rs.open "select id from 物品 where 物品名='" & wpname & "' and 擁有者="&info(9),conn
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品 (物品名,擁有者,類型,內力,體力,攻擊,防御,數量,銀兩,說明,堅固度,等級,sm,會員) values ('"&wpname&"',"&info(9)&",'"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &","& yin &",'"& say &"',"& jgd &","& dj &","& sm &","&aaao&")"
else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+" & sl & ",會員="&aaao&" where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示："& myname &"在自己的保險櫃中把["& wpname &"]取出{"& sl &"}個！');location.href = 'mybxg.asp';</script>"
response.end
%>