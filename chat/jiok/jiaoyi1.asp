<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
server_v1=Cstr(Request.ServerVariables("HTTP_REFERER"))
server_v2=Cstr(Request.ServerVariables("SERVER_NAME"))
if mid(server_v1,8,len(server_v2))<>server_v2 then
response.write "<br><br><center><table border=1 cellpadding=20 bordercolor=black bgcolor=#EEEEEE width=450>"
response.write "<tr><td style='font:9pt Verdana'>"
response.write "�A���檺���|���~�A�T��q���I�~������ƾڽФ��n�øӰѼơI"
response.write "</td></tr></table></center>"
response.end
end if

nickname=info(0)
nick1=info(9)
wupin=request("wupin")
sl=trim(request("sl"))
sl2=trim(request("sl2"))
money=request("money")
danwei=request("danwei")
danwei1="�Ȩ�"
n=Year(date())
y=Month(date())
r=Day(date())
s=Hour(time())
f=Minute(time())
m=Second(time())
if len(y)=1 then y="0" & y
if len(r)=1 then r="0" & r
if len(s)=1 then s="0" & s
if len(f)=1 then f="0" & f
if len(m)=1 then m="0" & m
sj=s & ":" & f & ":" & m
sj2=n & "-" & y & "-" & r & " " & sj
if sl>sl2  then
Response.Write "<script Language=Javascript>alert('���n�@���F�A�n�n���աI');window.close();</script>"
response.end
end if
if application("jiutian_jiaoyi")<>""  then
jian=split(application("jiutian_jiaoyi"),"|")
if DateDiff("s",jian(1),sj2)<60 then
Response.Write "<script Language=Javascript>alert('���e��������٦�" & 60-DateDiff("s",jian(1),sj2) & "��!');window.close();</script>"
response.end
else
Application.Lock
application("jiutian_jiaoyi")=nick1&"|"&sj2&"|"&money&"|"&wupin&"|"&sl&"|"&danwei&"|"&nickname
Application.UnLock
msg="�b�F�C����j��椤�A["&nickname&"]�[�ܨ�G���~<font color=FFFDAF>["&wupin&"]</font>�{�b���A�ƶq��<font color=red>"&sl&"</font>,�`�����欰�G<font color=red>"&money&""&danwei&"</font>, �j�a�ַm�ڡA<img src='jiok/LOC54.GIF'><input type=button value='�ʶR' onclick=javascript:window.open('jiok/jiaoyiok1.asp','a','width=300,height=250') style='font-size:9pt;background-color:FF9900;color:FFFFFF;border: 1 double'onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' >"

call jiaoyi()
end if
else
Application.Lock
application("jiutian_jiaoyi")=nick1&"|"&sj2&"|"&money&"|"&wupin&"|"&sl&"|"&danwei&"|"&nickname
Application.UnLock
msg="�b�F�C����j��椤�A["&nickname&"]�[�ܨ�G���~<font color=FFFDAF>["&wupin&"]</font>�{�b���A�ƶq��<font color=red>"&sl&"</font>,�`�����欰�G<font color=red>"&money&""&danwei&"</font>, �j�a�ַm�ڡA<img src='jiok/LOC54.GIF'><input type=button value='�ʶR' onclick=javascript:window.open('jiok/jiaoyiok1.asp','a','width=300,height=250') style='font-size:9pt;background-color:FF9900;color:FFFFFF;border: 1 double'onMouseOver=this.style.color='FFFF00' onMouseOut=this.style.color='FFFFFF' >"

call jiaoyi()
end if

sub jiaoyi()
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
sd(199)="<font color=FFFDAF>�i�F�C�����j" & msg &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script Language=Javascript>alert('�ާ@���\!');window.close();</script>"
response.end
end sub
%>