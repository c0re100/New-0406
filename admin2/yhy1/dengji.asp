<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if info(3)<3 then 
Response.Write "<script Language=Javascript>alert('���c���n�S�W�S�檺����p���I��I�I�I');window.close();</script>"
response.end
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT �ʧO,�y�O FROM �Τ� WHERE id="& info(9)
set rs=conn.execute(sql)
sex=rs("�ʧO")
meimao=rs("�y�O")
maiyin=meimao*100
if sex<>"�k" then 
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('�w~�A�ӳo�̰�����H���ۤp�j�@�I�I');window.close();</script>"
response.end
end if
if meimao<1000 then 
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('�A�]�Ӥ��F�a~�O�v�T�ڭ̥ͷN�I�h�h~�I');window.close();</script>"
response.end
end if
%><!--#include file="jiu.asp"--><% 
sql="select * from �Өk where �m�W='" & info(0) & "'"
set rs=connt.execute(sql)
if rs.bof or rs.eof then
sql="insert into �Өk(�m�W,������) values ('" & info(0) & "'," & meimao & ")"
connt.execute sql
conn.execute "update �Τ� set �Ȩ�=�Ȩ�+"&maiyin&" where id="&info(9)
connt.close
set connt=nothing
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('���ߡA�A�����������c���Өk�A�o��樭�Ȩ�"&maiyin&"��I���A�����F�A��ū���a�I');window.close();</script>"
response.end
else
connt.close
set connt=nothing
conn.close
set conn=nothing
set rs=nothing
Response.Write "<script Language=Javascript>alert('�A�w�g�O���c���Өk�F�A�٨Ӱ�����r�H');window.close();</script>"
response.end
end if
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFCFCF"
sd(197)="FFCFCF"
sd(198)="��"
sd(199)="<font color=FFCFCF>�i�樭�����j"&info(0)&"</font>���F�ͭp�A���o����ۤv�����c�A�o��樭��<font color=green>"&maiyin&"��</font>�I�u~~��~���~���ٲM�ڡI"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
%>