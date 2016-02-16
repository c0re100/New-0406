<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
men1=trim(Application("menpai"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select 幫派基金,適合 FROM 門派 WHERE 門派='" & men1 & "'",conn
if rs("幫派基金")<15000000  then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('因對方門派基金太少,暫時停止招收新弟子，請下次吧！');}</script>"
	Response.End
end if
shihe=trim(rs("適合"))
rs.close
rs.open "select 門派,身份 FROM 用戶 WHERE id="&info(9),conn
if rs("門派")="官府" or rs("門派")=men1 or rs("身份")="掌門" then
	Response.Write "<script language=JavaScript>{alert('看仔細了不要亂點！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if
if shihe<>info(1) and shihe<>"所有" then
	Response.Write "<script language=JavaScript>{alert('該門派隻招收"&shihe&"弟子！');}</script>"
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	Response.End
end if

conn.execute "update 門派 set 弟子數=弟子數-1 where 門派='" & info(5) & "'"
conn.execute "update 門派 set 弟子數=弟子數+1 where 門派='" & men1 & "'"
if rs("門派")<>"遊俠" then
conn.execute "update 用戶 set 保護=false,武功=100,內力=100,銀兩=0,存款=0 where id="&info(9)
jiamen="["&info(0)&"]跑到["&men1&"]的掌門面前說：我來啦,快給我錢啊,我要錢啊..."&info(0)&"背叛了原來的["&info(5)&"]門派,成功的加入["&men1&"]!"
else
conn.execute "update 用戶 set 銀兩=0,存款=0 where id="&info(9)
jiamen="["&info(0)&"]跑到["&men1&"]的掌門面前說：我要加入你們門派，快給我點錢啊，我快窮死了,"&info(0)&"成功的加入["&men1&"]!"
end if
conn.execute "update 門派 set 幫派基金=幫派基金-10000000 where 門派='" & men1 & "'"
conn.execute "update 用戶 set 門派='" & men1 & "', 身份='弟子',grade=1 where id="&info(9)
rs.close
	set rs=nothing
	conn.close
	set conn=nothing
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)="消息"
sd(195)=info(0)
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="對"
sd(199)="<font color=B0D0E0>【應招入派】</font>"&jiamen
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>
