<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
if info(2)<10   then Response.Redirect "../error.asp?id=900"
%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
cpass2=Request.Form("cpass2")
cpass22=int(Request.Form("cpass22"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")

conn.open Application("ljjh_usermdb")
sql="SELECT * FROM �Τ� where �m�W='"&cpass2&"'"
Set Rs=conn.Execute(sql)
if rs.EOF or rs.BOF then
Response.Write "<script language=javascript>alert('��p�A�ҭn�d�䪺�H�ڭ̧䤣��I�Ьd�ݬO�_���T�I');history.back();</script>"
else
aa=rs("����")
bb=(aa+cpass22)*(aa+cpass22)*50
sql="update �Τ� set ����=����+"&cpass22&",allvalue="&bb&" where �m�W='"&cpass2&"'"
Set Rs=conn.Execute(sql)
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�[�žާ@')"
conn.Execute(sql)
conn.close
set rs=nothing
set conn=nothing
Response.Write "<script language=javascript>alert('�Τ�G["&cpass2&"]�[["&cpass22&"]�ŭק令�\�I');history.back();</script>"
end if
%> 