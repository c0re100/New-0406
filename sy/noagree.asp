<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
if info(2)<10 then Response.Redirect "../error.asp?id=439"
result=request("result")
id=request("id")
sql="update �ӭ� set ���G='0' where id=" & id & ""
conn.execute sql
rs.open ("SELECT �Q�i,��i,�n�D,���D FROM �ӭ� WHERE ID=" & id),conn,0,1
bg=rs("�Q�i")
yg=rs("��i")
result=rs("�n�D")
mess=rs("���D")
newer="����H�h<font color=FFD7F4>" & yg & "</font>���i<font color=FFD7F4>" & bg & "</font>���D�G<font color=red>{"&mess&"}</font>,�n�D<font color=red>" & result & "</font>,�]�Ҿڤ����B�z�Ѥ��R���A���q�L�I"
rs.close
conn.close
set rs=nothing	
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFD7F4"
sd(197)="FFD7F4"
sd(198)="��"
sd(199)="<font color=FFD7F4>�i�ӭާP�M�j</font>"& newer 
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock

Response.Redirect "manage.asp"
%>