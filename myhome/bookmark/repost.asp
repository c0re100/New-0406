<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
%>

 
<!--#include file="../data1.asp"-->

<%
if trim(request.form("name"))="" or trim(request.form("type"))="" or trim(request.form("url"))="" then
%>
<script language="Vbscript">
msgbox"請填寫完整！",0,"提示"
history.back
</script>

<%
conn.close
Response.End
end if
%>
<!-- 判斷網址是否合格並處理 -->
<%
url=trim(request.form("url"))
if left(url,7)<>"http://" then
url="http://"&url
end if
%>

<!-- 寫入 -->
<%
rs.open "select * from bookmark where id="&request.form("id")&"",conn,1,3
rs.update "user",info(0)
rs.update "type",request.form("type")
rs.update "emptytype",request.form("emptytype")
rs.update "name",request.form("name")
rs.update "url",url
rs.update "open",request.form("open")
rs.update "datenow",ccdate(date)
rs.close
conn.close
%>
<script language="Vbscript">
msgbox"您已經成功地修改了書簽！",0,"提示"
history.back
</script>

<%
'=====================================================
'轉換為寫入數據庫的格式，去掉非法字符
'調用：putmeg(字符串)

function putmeg (inputmsg)
putmeg=replace(inputmsg,chr(13)&chr(10),"<br>")
putmeg=replace(putmeg," ","&nbsp;&nbsp;")
putmeg=replace(putmeg,"'","’")
putmeg=replace(putmeg,"<",".")

end function
'=====================================================


'=====================================================
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
'=====================================================

%>