<%




'聲明connection

dbname="driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("setup.mdb")
set conn=server.createobject("adodb.connection")
conn.open dbname
set rs=server.createobject("adodb.recordset")


'----------------------------------------------------------------------

'-----------------------------------------------------------------------
'得到文件擴展名的函數
'調用：getfileextname(文件全名)

function getFileExtName(fileName)
dim pos
pos=instrrev(filename,".")
if pos>0 then 
getFileExtName=mid(fileName,pos+1)
else
getFileExtName=""
end if
end function 
'判斷空字符串，如是，則用指定值填充
'調用：cnone(字符串,填充字符串)

function cnone(panduan,zhiding)
if trim(panduan)="" then
cnone=zhiding
else
cnone=panduan
end if
end function
'轉換日期型或字符日期型為全格式如：2000-01-01
'調用：ccdate(日期變量或日期格式字符串)

function ccdate(sdate)
temp=cdate(sdate)
if len(month(temp))=1 then
m="0"&month(temp)
else
m=month(temp)
end if

if len(day(temp))=1 then
d="0"&day(temp)
else
d=day(temp)
end if
ccdate=year(temp)&"-"&m&"-"&d
end function
'轉換時間字符串為全格式,如：01:01:01
'調用：cctime(字符串)

function cctime (stime)
if len(hour(stime))=1 then
h="0"&hour(stime)
else
h=hour(stime)
end if
if len(minute(stime))=1 then
m="0"&minute(stime)
else
m=minute(stime)
end if
if len(second(stime))=1 then
s="0"&second(stime)
else
s=second(stime)
end if
cctime=h&":"&m&":"&s
end function
'轉換為寫入數據庫的格式，去掉非法字符
'調用：putmeg(字符串)

function putmeg (inputmsg)
putmeg=replace(putmeg,"<",".")
putmeg=replace(inputmsg,chr(13)&chr(10),"<br>")
putmeg=replace(putmeg," ","&nbsp;&nbsp;")
putmeg=replace(putmeg,"'","’")


end function
'按概率返回true，精確至0.1%
'調用：prob(概率)
function prob(temprob)
randomize
if temprob>100 then
Response.Write( "Flush內部函數 prob 錯誤！<br>")
Response.Write("超出概率範圍")
Response.End
end if
pnumber=(rnd*1000)/10
if pnumber<temprob then
prob=true
else
prob=false
end if
end function
'還原數據庫寫入格式（putmsg的反操作）
'調用：getmsg（字符串或數據庫字段值）

function getmsg(inputmsg)
getmsg=replace(inputmsg,"’","'")
getmsg=replace(getmsg,"&nbsp;&nbsp;"," ")
getmsg=replace(getmsg,"<br>",chr(13)&chr(10))
end function
'判斷email地址的正確性
'調用：isemail(email地址字符串)，如正確返回true,不正確則返回false

Function IsEmail(Email)
ValidFlag = False
If (Email <> "") And (InStr(1, Email, "@") > 0) And (InStr(1, Email, ".") > 0) Then
atCount = 0
SpecialFlag = False
For atLoop = 1 To Len(Email)
atChr = Mid(Email, atLoop, 1)
If atChr = "@" Then atCount = atCount + 1
If (atChr >= Chr(32)) And (atChr <= Chr(44)) Then SpecialFlag = True
If (atChr = Chr(47)) Or (atChr = Chr(96)) Or (atChr >= Chr(123)) Then SpecialFlag = True
If (atChr >= Chr(58)) And (atChr <= Chr(63)) Then SpecialFlag = True
If (atChr >= Chr(91)) And (atChr <= Chr(94)) Then SpecialFlag = True
Next
If (atCount = 1) And (SpecialFlag = False) Then
BadFlag = False
tAry1 = Split(Email, "@")
UserName = tAry1(0)
DomainName = tAry1(1)
If (UserName = "") Or (DomainName = "") Then BadFlag = True
If Mid(DomainName, 1, 1) = "." then BadFlag = True
If Mid(DomainName, Len(DomainName), 1) = "." then BadFlag = True
ValidFlag = True
End If
End If
If BadFlag = True Then ValidFlag = False
IsEmail = ValidFlag
End Function
'=====================================================

'截取一定長度的字符串，其餘補" ..."
'調用：cutstr(字符串，截取長度)

function cutstr(tempstr,tempwid)


if instr(tempstr,"&nbsp;") then
tempstr=replace(tempstr,"&nbsp;"," ")
end if
if instr(tempstr,"<br>") then
tempstr=replace(tempstr,"<br>"," ")
end if

if len(tempstr)>tempwid then
cutstr=left(tempstr,tempwid)&" ..."
else
cutstr=tempstr
end if
end function
'=====================================================

%>
