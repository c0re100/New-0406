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
	Response.Write "<script language=JavaScript>{alert('ぃ秈︽巨秈︽巨ゲ斗秈册ぱ');window.close();}</script>"
	Response.End 
end if
http = Request.ServerVariables("HTTP_REFERER") 
if InStr(http,"hcjs/es")=0 then 
Response.Write "<script language=javascript>{alert('癸ぃ癬祘┶荡眤巨\n     絋﹚');parent.history.go(-1);}</script>" 
Response.End 
end if
sl=abs(int(request.querystring("sl")))
id=lcase(trim(request.querystring("id")))
if InStr(id,"or")<>0 or InStr(id,"=")<>0 or InStr(id,"`")<>0 or InStr(id,"'")<>0 or InStr(id," ")<>0 or InStr(id," ")<>0 or InStr(id,"'")<>0 or InStr(id,chr(34))<>0 or InStr(id,"\")<>0 or InStr(id,",")<>0 or InStr(id,"<")<>0 or InStr(id,">")<>0 then Response.Redirect "../../error.asp?id=54"
myname=info(0)
money=abs(Request.form("money"))
zy=LCase(trim(Request.form("zy")))
zy=replace(zy,"'","")
zy=replace(zy,chr(34),"")
zy=Replace(zy,"<","")
zy=Replace(zy,">","")
zy=Replace(zy,"\x3c","")
zy=Replace(zy,"\x3e","")
zy=Replace(zy,"\074","")
zy=Replace(zy,"\74","")
zy=Replace(zy,"\75","")
zy=Replace(zy,"\76","")
zy=Replace(zy,"&lt","")
zy=Replace(zy,"&gt","")
zy=Replace(zy,"\076","")
badstr="射精|奸|去死|吃屎|你妈|你娘|日你|尻|操你|干死你|王八|逼|傻B|贱人|狗娘|婊子|表子|靠你|叉你|叉死|插你|插死|干你|干死|日死|鸡巴|睾丸|死去 |爬你达来蛋|撅你达来蛋|死你达来蛋|包皮|龟头|屄|赑|妣|肏|奶子|尻|屌|作爱|做爱|床上|抱抱|鸡八|处女|打炮|十八摸|你爷|你爸|我儿|操你|妈|逼|asp|com|net|www|xajh|202|61|jh|江湖|or|261|网管|掌门"
bad=split(badstr,"|")
for i=0 to UBound(bad)
zy=Replace(zy,bad(i),"**")
next
if instr(zy,"or")<>0 or instr(zy,"'")<>0 then Response.Redirect "../../error.asp?id=54"
if zy="" then
	Response.Write "<script Language=Javascript>alert('矗ボもユぃΘ﹎籔弧ぃ');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 珇,摸,ず,砰,ň眘,ю阑,蝗ㄢ,弧,绊㏕,单,sm from 珇 where id=" & id & " and 摸<>'' and 计秖>="&sl&" and 局Τ="&info(9),conn
if rs.eof or rs.bof then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script Language=Javascript>alert('矗ボ巨ぃΘ⊿Τ硂妓珇');location.href = 'javascript:history.go(-1)';</script>"
	response.end
end if
wpname=rs("珇")
lx=rs("摸")
nl=rs("ず")
tl=rs("砰")
fy=rs("ň眘")
gj=rs("ю阑")
yin=rs("蝗ㄢ")
say=rs("弧")
sm=rs("sm")
dj=rs("单")
jgd=rs("绊㏕")
if money>9999999 then
 money=9999999
end if
conn.execute "update 珇 set 计秖=计秖-"& sl &",穦="&aaao&" where id=" & id
conn.execute "insert into ユカ初 (珇,局Τ,よΑ,癸钩,摸,ず,砰,ю阑,ň眘,计秖,弧,も基,丁,蝗ㄢ,sm,绊㏕,单,say,穦) values ('"&wpname&"',"& info(9) &",'も','產','"&lx& "',"& nl &","& tl &","& gj &","& fy &","& sl &",'"& say &"',"& money &",now(),"& yin &","& sm &","& jgd &","& dj &",'"& zy &"',"&aaao&")"
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Write "<script Language=Javascript>alert('矗ボ"& myname &"芥も珇:"& wpname & sl &"矗ユЧΘ单ユ');location.href = 'wupin.asp';</script>"
response.end
%>
