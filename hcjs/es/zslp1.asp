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
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
myname=info(0)
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 對像,物品名,類型,內力,體力,防御,攻擊,銀兩,數量,二手價,擁有者,說明,sm,堅固度,等級 from 交易市場 where id=" & id & " and 數量>0",conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：情意禮品屋並沒有這種物品！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
duixiang=rs("對像")
wpname=rs("物品名")
lx=rs("類型")
nl=rs("內力")
tl=rs("體力")
fy=rs("防御")
gj=rs("攻擊")
yin=rs("銀兩")
sl=rs("數量")
say=rs("說明")
sm=rs("sm")
dj=rs("等級")
jgd=rs("堅固度")
esmoney=rs("二手價")
toname=rs("擁有者")
if duixiang<>myname then
	conn.close
	set rs=nothing
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('提示：呵。。笑話，想作什麼呀，這個也不是送給你的！');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
conn.execute "delete * from 交易市場 where  id="&id
rs.close
rs.open "select id from 物品 where 物品名='" & wpname & "' and 擁有者="&info(9)
If Rs.Bof OR Rs.Eof then
	conn.execute "insert into 物品 (物品名,擁有者,類型,內力,體力,攻擊,防御,數量,銀兩,說明,堅固度,等級,sm,會員) values ('"&wpname&"',"&info(9)&",'"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &","& yin &",'"& say &"',"& jgd &"","& dj &","& sm &","&aaao&")"
else
	id=rs("id")
	conn.execute "update 物品 set 數量=數量+" & sl & ",會員="&aaao&" where id="&id
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('提示："& myname &"在情意禮品屋領取["& wpname &"]成功！');location.href = 'zslp.asp';</script>"
response.end
%>


