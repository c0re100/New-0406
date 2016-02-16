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
conn.Execute "update ゅセ set 簎笆約='"& gg1 &"',穦纒='"& hy13 &"',穦='"& hy14 &"',='"& hy15 &"',穦る1="& hy1 &",穦る2="& hy2 &",穦る3="& hy3 &",穦る4="& hy4 &",穦る5="& hy17 &",穦る6="& hy18 &",穦る7="& hy19 &",穦る8="& hy20 &",穦る9="& hy21 &",穦る10="& hy22 &" where id="&id
conn.Execute "update ゅセ set 穦单1="& hy5 &",穦单2="& hy6 &",穦单3="& hy7 &",穦单4="& hy8 &",穦单5="& hy23 &",穦单6="& hy24 &",穦单7="& hy25 &",穦单8="& hy26 &",穦单9="& hy27 &",穦单10="& hy28 &" where id="&id
conn.Execute "update ゅセ set 獶穦獁翴="&hy91&",穦獁翴1="& hy9 &",穦獁翴2="& hy10 &",穦獁翴3="& hy11 &",穦獁翴4="& hy12 &",穦獁翴5="& hy29 &",穦獁翴6="& hy30 &",穦獁翴7="& hy31 &",穦獁翴8="& hy32 &",穦獁翴9="& hy33 &",穦獁翴10="& hy34 &" where id="&id
conn.Execute "update ゅセ set PK='"&hy16&"' where id="&id
sql="insert into 恨瞶癘魁 (﹎,丁,ip,癘魁) values ('"&info(0)&"',now(),'"&info(15)&"','約э巨')"
conn.Execute(sql)
	set rs=nothing	
	conn.close
	set conn=nothing
Response.Redirect "../ok.asp?id=700"
%> 