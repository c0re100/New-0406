<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
db=request.querystring("db")
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
	Response.End 
end if
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W FROM �Τ� WHERE id="&id,conn
peiou=rs("�m�W")
if info(0)<>trim(Application("ljjh_kissly")) then
	rs.close
	set rs=nothing
	conn.close
	set conn=nothing
	Response.Write "<script language=JavaScript>{alert('�A�Q�@����A�H�a�]�S�����M�A��աI');}</script>"
	Response.End
end if
if db-Application("ljjh_bingwen")=-2 or db-Application("ljjh_bingwen")=1 then
	Response.Write "<script language=JavaScript>{alert('�A��F�A�˾`�ʪ��I�u~~~');}</script>"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�-1000000 where id="&info(9)
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+2000000 where id="&id
        hunyin="�˾`�G["& info(0) &"]��ո̿鵹�F{"& peiou &"}�A100�U��ժ�᪺�Ȥl���F�I"
end if
if db-Application("ljjh_bingwen")=-1 or db-Application("ljjh_bingwen")=2 then
        Response.Write "<script language=JavaScript>{alert('�A�ӤF�A�����A100�U��@~~~');}</script>"
	conn.execute "update �Τ� set �Ȩ�=�Ȩ�+1000000 where id="&info(9)
	hunyin="���ߡG["& info(0) &"]��ո�Ĺ�F{"& peiou &"}100�U��Ȥl�I"
end if
if db=Application("ljjh_bingwen") then
	Response.Write "<script language=JavaScript>{alert('���F�A���s�ӧa�I');}</script>"
        conn.execute "update �Τ� set �Ȩ�=�Ȩ�+1000000 where id="&id
	hunyin="�����G["& info(0) &"]��ո̩M{"& peiou &"}�����F�A���j���Y�����ѦӦX�C"
end if
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
sd(194)="����"
sd(195)="�j�a"
sd(196)="FFFDAF"
sd(197)="FFFDAF"
sd(198)="��"
sd(199)="<font color=red>�i��������j</font>"&hunyin
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application("ljjh_kissly")=""
Application("ljjh_bingwen")=""
Application.UnLock
%>
