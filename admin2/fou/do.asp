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
sql="SELECT 銀兩,操作時間 FROM 用戶 where id="&info(9)
rs.open sql,conn,1,1
sj=DateDiff("s",rs("操作時間"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('請你等上["& ss &"]秒,再操作！');location.href = 'index.asp';}</script>"
	Response.End
end if
yin1=("銀兩")
jiu=request.form("jiu")
	select case jiu
           case "lao"
              tili=40
              yin=200000
           case "wu"
              tili=60
              yin=300000
           case "du"
              tili=80
              yin=400000
           case "mao"
              tili=100
              yin=500000
          end select
		if yin1<yin then
			mess="<font color=FFCFCF>[" & info(0) & "]</font>，你沒錢也想來充大款啊，滾吧你!!"
else
        randomize timer
        r=int(rnd*3)+1
if r=1 or r=3 then
                       	sql="update 用戶 set 銀兩=銀兩-" & yin & ",道德=道德+"& tili &",操作時間=now() where id="&info(9)
			conn.execute sql
			mess="<font color=FFCFCF>[" & info(0) & "]</font>從相國寺出來後，覺得自己得益非淺，同時也悟出不少道理!<font color=red>道德增加"& tili &"</font>!!!"
else
sql="update 用戶 set 銀兩=銀兩-50000,魅力=魅力-40,操作時間=now() where id="&info(9)
mess="<font color=FFCFCF>[" & info(0) & "]</font>本想去相國寺做善事，不料半路遇到強盜，被搶去50000兩銀子!<font color=red>魅力減了40</font>!!!"
end if
end if
rs.close
conn.close
set conn=nothing
set rs=nothing
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
sd(195)="大家"
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="對"
sd(199)="<font color=FFCFCF>【小道消息】</font><font color=red>"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{alert('操作成功，確定返回！');location.href = 'indexdo.asp';}</script>"
%>