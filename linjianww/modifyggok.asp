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
<%Response.Expires=0
if session("ljjh_adminok")<>true then Response.Redirect "../chat/readonly/bomb.htm"
nickname=info(0)
grade=Int(info(2))
if info(2)<10   then Response.Redirect "../error.asp?id=900"
id=trim(Request.Form("gg"))
gg1=trim(Request.Form("gg1"))
hy1=abs(trim(Request.Form("hy1")))
hy2=abs(trim(Request.Form("hy2")))
hy3=abs(trim(Request.Form("hy3")))
hy4=abs(trim(Request.Form("hy4")))
hy5=abs(trim(Request.Form("hy5")))
hy6=abs(trim(Request.Form("hy6")))
hy7=abs(trim(Request.Form("hy7")))
hy8=abs(trim(Request.Form("hy8")))
hy9=abs(trim(Request.Form("hy9")))
hy91=abs(trim(Request.Form("hy91")))
hy10=abs(trim(Request.Form("hy10")))
hy11=abs(trim(Request.Form("hy11")))
hy12=abs(trim(Request.Form("hy12")))
hy17=abs(trim(Request.Form("hy17")))
hy18=abs(trim(Request.Form("hy18")))
hy19=abs(trim(Request.Form("hy19")))
hy20=abs(trim(Request.Form("hy20")))
hy21=abs(trim(Request.Form("hy21")))
hy22=abs(trim(Request.Form("hy22")))
hy23=abs(trim(Request.Form("hy23")))
hy24=abs(trim(Request.Form("hy24")))
hy25=abs(trim(Request.Form("hy25")))
hy26=abs(trim(Request.Form("hy26")))
hy27=abs(trim(Request.Form("hy27")))
hy28=abs(trim(Request.Form("hy28")))
hy29=abs(trim(Request.Form("hy29")))
hy30=abs(trim(Request.Form("hy30")))
hy31=abs(trim(Request.Form("hy31")))
hy32=abs(trim(Request.Form("hy32")))
hy33=abs(trim(Request.Form("hy33")))
hy34=abs(trim(Request.Form("hy34")))
hy13=trim(Request.Form("hy13"))
hy14=trim(Request.Form("hy14"))
hy15=trim(Request.Form("hy15"))
hy16=trim(Request.Form("hy16"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
conn.Execute "update �奻 set �u�ʼs�i='"& gg1 &"',�|���s�d='"& hy13 &"',�|���a�}='"& hy14 &"',�b�~�I='"& hy15 &"',�|����I1="& hy1 &",�|����I2="& hy2 &",�|����I3="& hy3 &",�|����I4="& hy4 &",�|����I5="& hy17 &",�|����I6="& hy18 &",�|����I7="& hy19 &",�|����I8="& hy20 &",�|����I9="& hy21 &",�|����I10="& hy22 &" where id="&id
conn.Execute "update �奻 set �|������1="& hy5 &",�|������2="& hy6 &",�|������3="& hy7 &",�|������4="& hy8 &",�|������5="& hy23 &",�|������6="& hy24 &",�|������7="& hy25 &",�|������8="& hy26 &",�|������9="& hy27 &",�|������10="& hy28 &" where id="&id
conn.Execute "update �奻 set �D�|���w�I="&hy91&",�|���w�I1="& hy9 &",�|���w�I2="& hy10 &",�|���w�I3="& hy11 &",�|���w�I4="& hy12 &",�|���w�I5="& hy29 &",�|���w�I6="& hy30 &",�|���w�I7="& hy31 &",�|���w�I8="& hy32 &",�|���w�I9="& hy33 &",�|���w�I10="& hy34 &" where id="&id
conn.Execute "update �奻 set PK��='"&hy16&"' where id="&id
sql="insert into �޲z�O�� (�m�W,�ɶ�,ip,�O��) values ('"&info(0)&"',now(),'"&info(15)&"','�s�i�ק�ާ@')"
conn.Execute(sql)
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Redirect "../ok.asp?id=700"
%> 