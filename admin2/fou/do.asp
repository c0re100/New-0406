<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT �Ȩ�,�ާ@�ɶ� FROM �Τ� where id="&info(9)
rs.open sql,conn,1,1
sj=DateDiff("s",rs("�ާ@�ɶ�"),now())
if sj<8 then
	rs.close
	set rs=nothing	
	conn.close
	set conn=nothing
	ss=8-sj
	Response.Write "<script language=JavaScript>{alert('�ЧA���W["& ss &"]��,�A�ާ@�I');location.href = 'index.asp';}</script>"
	Response.End
end if
yin1=("�Ȩ�")
jiu=request.form("jiu")
	select case jiu
           case "lao"
              tili=40
              yin=200000
           case "wu"
              tili=60
              yin=300000
           case "du"
              tili=80
              yin=400000
           case "mao"
              tili=100
              yin=500000
          end select
		if yin1<yin then
			mess="<font color=FFCFCF>[" & info(0) & "]</font>�A�A�S���]�Q�ӥR�j�ڰڡA�u�a�A!!"
else
        randomize timer
        r=int(rnd*3)+1
if r=1 or r=3 then
                       	sql="update �Τ� set �Ȩ�=�Ȩ�-" & yin & ",�D�w=�D�w+"& tili &",�ާ@�ɶ�=now() where id="&info(9)
			conn.execute sql
			mess="<font color=FFCFCF>[" & info(0) & "]</font>�q�۰�x�X�ӫ�Aı�o�ۤv�o�q�D�L�A�P�ɤ]���X���ֹD�z!<font color=red>�D�w�W�["& tili &"</font>!!!"
else
sql="update �Τ� set �Ȩ�=�Ȩ�-50000,�y�O=�y�O-40,�ާ@�ɶ�=now() where id="&info(9)
mess="<font color=FFCFCF>[" & info(0) & "]</font>���Q�h�۰�x�����ơA���ƥb���J��j�s�A�Q�m�h50000��Ȥl!<font color=red>�y�O��F40</font>!!!"
end if
end if
rs.close
conn.close
set conn=nothing
set rs=nothing
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
sd(199)="<font color=FFCFCF>�i�p�D�����j</font><font color=red>"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
Response.Write "<script language=JavaScript>{alert('�ާ@���\�A�T�w��^�I');location.href = 'indexdo.asp';}</script>"
%>