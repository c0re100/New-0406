<%




'�n��connection

dbname="driver={Microsoft Access Driver (*.mdb)};DBQ="&Server.MapPath("setup.mdb")
set conn=server.createobject("adodb.connection")
conn.open dbname
set rs=server.createobject("adodb.recordset")


'----------------------------------------------------------------------

'-----------------------------------------------------------------------
'�o�����X�i�W�����
'�եΡGgetfileextname(�����W)

function getFileExtName(fileName)
dim pos
pos=instrrev(filename,".")
if pos>0 then 
getFileExtName=mid(fileName,pos+1)
else
getFileExtName=""
end if
end function 
'�P�_�Ŧr�Ŧ�A�p�O�A�h�Ϋ��w�ȶ�R
'�եΡGcnone(�r�Ŧ�,��R�r�Ŧ�)

function cnone(panduan,zhiding)
if trim(panduan)="" then
cnone=zhiding
else
cnone=panduan
end if
end function
'�ഫ������Φr�Ť���������榡�p�G2000-01-01
'�եΡGccdate(����ܶq�Τ���榡�r�Ŧ�)

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
'�ഫ�ɶ��r�Ŧꬰ���榡,�p�G01:01:01
'�եΡGcctime(�r�Ŧ�)

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
'�ഫ���g�J�ƾڮw���榡�A�h���D�k�r��
'�եΡGputmeg(�r�Ŧ�)

function putmeg (inputmsg)
putmeg=replace(putmeg,"<",".")
putmeg=replace(inputmsg,chr(13)&chr(10),"<br>")
putmeg=replace(putmeg," ","&nbsp;&nbsp;")
putmeg=replace(putmeg,"'","��")


end function
'�����v��^true�A��T��0.1%
'�եΡGprob(���v)
function prob(temprob)
randomize
if temprob>100 then
Response.Write( "Flush������� prob ���~�I<br>")
Response.Write("�W�X���v�d��")
Response.End
end if
pnumber=(rnd*1000)/10
if pnumber<temprob then
prob=true
else
prob=false
end if
end function
'�٭�ƾڮw�g�J�榡�]putmsg���Ͼާ@�^
'�եΡGgetmsg�]�r�Ŧ�μƾڮw�r�q�ȡ^

function getmsg(inputmsg)
getmsg=replace(inputmsg,"��","'")
getmsg=replace(getmsg,"&nbsp;&nbsp;"," ")
getmsg=replace(getmsg,"<br>",chr(13)&chr(10))
end function
'�P�_email�a�}�����T��
'�եΡGisemail(email�a�}�r�Ŧ�)�A�p���T��^true,�����T�h��^false

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

'�I���@�w���ת��r�Ŧ�A��l��" ..."
'�եΡGcutstr(�r�Ŧ�A�I������)

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
