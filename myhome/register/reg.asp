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

<!--#include file="data2.asp"-->
<!--#INCLUDE file="check.asp"-->
<%
c1=check(request.form("name"),"姓名")
c2=check(request.form("email"),"電子郵箱")
c1=check(request.form("homepage"),"個人主頁")
c2=check(request.form("address"),"家庭地址")
c1=check(request.form("pc"),"郵政編碼")
c2=check(request.form("tel"),"聯繫電話")
c1=check(request.form("job"),"職業")
c2=check(request.form("idcard"),"身份證號")
c1=check(request.form("corporation"),"工作單位")
c2=check(request.form("weight"),"身高")
c1=check(request.form("nation"),"民族")
c2=check(request.form("icq"),"icq")
c1=check(request.form("oicq"),"oicq")
c2=check(request.form("pager"),"尋呼")
c1=check(request.form("mobilephone"),"移動電話")
c1=check(request.form("preference"),"個人介紹")
c2=check(request.form("introduction"),"個人愛好")
%>
<html>

<head>
<%name=request.form("name")
sex =request.form("sex")
birthyear=request.form("birthyear")
birthmonth=request.form("birthmonth")
birthdate=request.form("birthdate")
email =request.form("email")
homepage=request.form("homepage")
address=request.form("address")
pc =request.form("pc")
tel=request.form("tel")
job=request.form("job")
idcard =request.form("idcard")
corporation=request.form("corporation")
salary=request.form("salary")
mode1=request.form("mode1")
mode2=request.form("mode2")
mode3=request.form("mode3")
stature=request.form("stature")
marriage=request.form("marriage")
weight=request.form("weight")
city=request.form("city")
nation=request.form("nation")
education=request.form("education")
icq=request.form("icq")
health=request.form("health")
oicq=request.form("oicq")
pager=request.form("pager")
mobilephone=request.form("mobilephone")
preference=request.form("preference")
introduction=request.form("introduction")
open=request.form("open")
rs.open"select * from  userinfo where user='"&info(0)&"'",conn,1,3
rs.update "name",name
rs.update "sex",sex
rs.update "birthday",birthyear&"年"&birthmonth&"月"&birthdate&"日"
rs.update "email",email
rs.update "homepage",homepage
rs.update "address",address
rs.update "pc",pc
rs.update "tel", tel
rs.update "job",job
rs.update "idcard",idcard
rs.update "corporation",corporation
rs.update "salary",salary
rs.update "stature",stature
rs.update "marriage",marriage
rs.update "weight",weight
rs.update "city",city
rs.update "nation",nation
rs.update "education",education
rs.update "icq",icq
rs.update "health",health
rs.update "oicq",oicq
rs.update "pager",pager
rs.update "mobilephone",mobilephone
rs.update "preference",preference
rs.update "introduction",introduction
rs.update "open",open
rs.close
Set Conn=Server.CreateObject("ADODB.Connection")
Connstr="DBQ="+server.mappath("../../linjiankkk/asz.asa")+";DefaultDir='';DRIVER={Microsoft Access Driver (*.mdb)};DriverId=25;FIL=MS Access;ImplicitCommitSync=Yes;MaxBufferSize=512;MaxScanRows=8;PageTimeout=5;SafeTransactions=0;Threads=3;UserCommitSync=Yes;"
conn.open ljjh_mdb(0)
set rs=server.createobject("adodb.recordset")
rs.open"select * from  用戶 where 姓名='"&info(0)&"'",conn,1,3
rs.update "信箱",email
rs.update "性別",sex
response.write"已經成功修改資料！"
rs.close

%>

<title>真實資料</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>

<a href="javascript:history.back(1)">返回</a>

</body>

</html>