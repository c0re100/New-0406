<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
session("ljjh_jm")=session("ljjh_jm")+1
function chuser(u)
dim filter,xx,usernameenable,su
for i=1 to len(u)
su=mid(u,i,1)
xx=asc(su)
zhengchu = -1*xx \ 256
yushu = -1*xx mod 256
if (xx>122 or (xx>57 and xx<97) or (xx<-10241 and xx>-10247) or yushu=129 or yushu>192 or (yushu<2 and yushu>-1) or (((zhengchu>1 and zhengchu<8) or (zhengchu>79 and zhengchu<86)) and yushu<96 ) or (xx>-352 and xx<48) or (xx<-22016 and xx>-24321) or (xx<-32448)) then
chuser=true
exit function
end if
next
chuser=false
end function
regjm=Request.form("regjm")
regjm1=Request.form("regjm1")
if regjm<>regjm1 then
%>
<script language=vbscript>
alert ("�A��J���{�ҽX�����T�A���ӿ�J:<%=regjm%>")
location.href = "javascript:history.back()"
</script>
<%
response.end
end if
name=lcase(trim(request.form("name")))
pass=request.form("pass")
name1=lcase(trim(request.form("name1")))
for i=1 to len(name1)
usernamechr=mid(name1,i,1)
if asc(usernamechr)>0 then Response.Redirect "../error.asp?id=120"
next
if name1="�L" or name="�L" or name="���w" or name1="���w" then Response.Redirect "../error.asp?id=130"
if left(name1,3)="%20" OR InStr(name1,"=")<>0 or InStr(name1,"`")<>0 or InStr(name1,"'")<>0 or InStr(name1," ")<>0 or InStr(name1," ")<>0 or InStr(name1,"'")<>0 or InStr(name1,chr(34))<>0 or InStr(name1,"\")<>0 or InStr(name1,",")<>0 or InStr(name1,"<")<>0 or InStr(name1,">")<>0 then Response.Redirect "../error.asp?id=120"
if chuser(name1) then Response.Redirect "../error.asp?id=120"
if instr(name,"or")<>0 or instr(name,"'")<>0 or instr(name,"|")<>0 or instr(name," ")<>0 then Response.Redirect "../error.asp?id=120"
if instr(name1,"or")<>0 or instr(name1,"'")<>0 or instr(name1,"|")<>0 or instr(name1," ")<>0 then Response.Redirect "../error.asp?id=120"
'if Instr(name1,"�B�ͦW�r")>0 or Instr(name1,Application("sjjh_automanname"))>0 or Instr(name1,"���մ���")>0 or Instr(name1,"����")>0 or Instr(name1,"����")>0 or Instr(name1,"�L")>0 or Instr(name1,"����")>0 or Instr(name1,"��")>0 or Instr(name1,"��")>0 or Instr(name1,"�j�a")>0 or Instr(name1,"��")>0 then Response.Redirect "../error.asp?id=130"
if name="" or name1="" or pass="" then Response.Redirect "../error.asp?id=56"
'if Instr(LCase(Application("sjjh_useronlinename"))," "&LCase(name)&" ")<>0 then Response.Redirect "../error.asp?id=61"
temppass=StrReverse(left(pass&"godxtll,./",10))
templen=len(pass)
mmpassword=""
for j=1 to 10
mmpassword=mmpassword+chr(asc(mid(temppass,j,1))-templen+int(j*1.1))
next
pass=replace(mmpassword,"'","B")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
'����Τ�
sql="SELECT id FROM �Τ� WHERE �m�W='" & name1 & "'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  Response.Redirect "../error.asp?id=62"
sql="SELECT �Ȩ�,����,����,grade,�|������,�ʧO,�t�� FROM �Τ� WHERE �m�W='" & name & "' and �K�X='" & pass & "'"
Set Rs=conn.Execute(sql)
If Rs.Bof OR Rs.Eof Then Response.Redirect "../error.asp?id=63"
if rs("�Ȩ�")<1000000 then Response.Redirect "../error.asp?id=457"
if rs("����")="�x��" or rs("����")="�x��" or rs("grade")>=6 then Response.Redirect "../error.asp?id=64"
if rs("�|������")>0 then Response.Redirect "../error.asp?id=65"
sex=rs("�ʧO")
bl=rs("�t��")
sql="update �Τ� set �m�W='" & name1 & "',�Ȩ�=�Ȩ�-1000000 where �m�W='" & name & "'"
conn.Execute sql
sql="update �Τ� set �v��='" & name1 & "' where �v��='" & name & "'"
conn.Execute sql
if bl<>"�L" and bl<>"" then
sql="update �Τ� set �t��='" & name1 & "' where �m�W='" & bl & "'"
conn.Execute sql
if sex="�k" then
conn.execute "update �X��� set �m�W�k='"& name1 &"' where �m�W�k='" & name & "'"
else
conn.execute "update �X��� set �m�W�k='"& name1 &"' where �m�W�k='" & name & "'"
end if
end if
conn.execute "update �U�� set �U�ڤH='"& name1 &"' where �U�ڤH='" & name & "'"
conn.execute "insert into �ާ@�O��(�ɶ�,�m�W1,�m�W2,�覡,�ƾ�) values (now(),'"& name &"','"& name1 &"','�ާ@','�ﴫ�W�r�I')"
set rs=nothing
conn.close
set conn=nothing
'BBS�׽�
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
bbs="DBQ="+server.mappath("../bbs/bbs.asa")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open bbs
'����Τ�
sql="SELECT * FROM book WHERE username='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "update book set username='"&name1&"' where username='"&name&"'"
end if
set rs=nothing
conn.close
set conn=nothing
'�O��
'BBS�׽�
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
gupiao="DBQ="+server.mappath("../gupiao/gp.asp")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};"
conn.open gupiao
'����Τ�
sql="SELECT * FROM �j�� WHERE �b��='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "update �j�� set �b��='"&name1&"' where �b��='"&name&"'"
end if
sql="SELECT * FROM �Ȥ� WHERE �b��='"&name&"'"
Set Rs=conn.Execute(sql)
If not rs.bof and not rs.eof  Then  
conn.execute "update �Ȥ� set �b��='"&name1&"' where �b��='"&name&"'"
end if
'�O��
set rs=nothing
conn.close
set conn=nothing
'�O��
Response.Redirect "../ok.asp?id=111"
%>

 