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
c1=check(request.form("name"),"�m�W")
c2=check(request.form("email"),"�q�l�l�c")
c1=check(request.form("homepage"),"�ӤH�D��")
c2=check(request.form("address"),"�a�x�a�}")
c1=check(request.form("pc"),"�l�F�s�X")
c2=check(request.form("tel"),"�pô�q��")
c1=check(request.form("job"),"¾�~")
c2=check(request.form("idcard"),"�����Ҹ�")
c1=check(request.form("corporation"),"�u�@���")
c2=check(request.form("weight"),"����")
c1=check(request.form("nation"),"����")
c2=check(request.form("icq"),"icq")
c1=check(request.form("oicq"),"oicq")
c2=check(request.form("pager"),"�M�I")
c1=check(request.form("mobilephone"),"���ʹq��")
c1=check(request.form("preference"),"�ӤH����")
c2=check(request.form("introduction"),"�ӤH�R�n")
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
rs.update "birthday",birthyear&"�~"&birthmonth&"��"&birthdate&"��"
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
rs.open"select * from  �Τ� where �m�W='"&info(0)&"'",conn,1,3
rs.update "�H�c",email
rs.update "�ʧO",sex
response.write"�w�g���\�ק��ơI"
rs.close

%>

<title>�u����</title>
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>

<body>

<a href="javascript:history.back(1)">��^</a>

</body>

</html>