<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
name1=Request.form("name12")
name1=trim(name1)
name1=server.HTMLEncode(name1)
name2=Request.form("name22")
name2=trim(name2)
name2=server.HTMLEncode(name2)
pass=request.form("pass2")
pass=trim(pass)
mess=Request.form("mess2")
mess=trim(mess)
mess=server.HTMLEncode(mess)
if name1="" or name2="" then
Response.Redirect "../error.asp?id=433"
elseif pass="" then
Response.Redirect "../error.asp?id=432"
elseif name1=name2 then
Response.Redirect "../error.asp?id=434"
else

temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")

%><head>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<%
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(0)
'校驗用戶 魅力大於100，錢大於1000
sql="SELECT 配偶,性別,銀兩 FROM 用戶 WHERE 姓名='" & name1 & "' and 密碼='" & pass & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then
conn.close
Response.Redirect "../error.asp?id=431"
response.end
else
if rs("配偶")<>"無" then
conn.close
Response.Redirect "../error.asp?id=430"
response.end
end if
xb1=trim(rs("性別"))
if rs("銀兩")>=1000 then
rs.close
set rs=nothing
set rs=conn.execute("SELECT 性別 FROM 用戶 WHERE trim(姓名)='" & name2 & "'")
if not (rs.bof or rs.eof) then
xb2=trim(rs("性別"))
if xb1<>xb2 then
sql="update 用戶 set 銀兩=銀兩-1000 where 姓名='" & name1& "'"
conn.execute sql
sz = "'" & name1 & "','" & name2 & "','" & mess & "', now()"
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
ljjh_mdb= split(Application("ljjh_mdb"),"|")
conn.open ljjh_mdb(2)
into_db = "INSERT INTO 月老 (姓名1, 姓名2, 說明, 時間) VALUES(" & sz & ")"
conn.Execute(into_db)
sql="delete * from 月老 where now()-時間>10"
conn.execute sql
Response.Redirect "../ok.asp?id=600"
else
Response.Redirect "../error.asp?id=435" %>
<script language=vbscript>
MsgBox "開什麼玩笑，你們倆性別一樣啊！江湖裡是不允許同性戀的。"
location.href = "../jh.asp"
</script>

<%                                        end if
else %>
<script language=vbscript>
MsgBox "江湖中沒有你要娶的人啊？搞錯了吧！"
location.href = "../jh.asp"
</script>

<%               end if
else %>
<script language=vbscript>
MsgBox "你哪裡有1000兩啊，沒錢還想娶媳婦？做夢去吧！"
location.href = "../jh.asp"
</script>

<%		end if
end if
conn.close
end if
%>
<script language=vbscript>
MsgBox "<%=message%>"
location.href = "../jh.asp"
</script>

 