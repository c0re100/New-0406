<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
useronlinename=Application("ljjh_useronlinename"&session("nowinroom"))
if Instr(LCase(useronlinename),LCase(" "&info(0)&" "))=0 then 
Session("ljjh_inthechat")="0" 
Response.Redirect "chaterr.asp?id=001" 
end if 
grade=Int(info(2))
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
if DateDiff("s",Session("ljjh_lasttime"),sj2)<2 then
Response.Write "<script language=JavaScript>{alert('���ܺC�C���A�O�O�ۮ@�I');}</script>"
Response.End
end if
t="..."
'�O�_�Q�I��
Application.Lock
if Instr(LCase(application("ljjh_dianxuename")),LCase(info(0)))>0 then
dianxue=split(application("ljjh_dianxuename"),";")
dxdx=0
for x=0 to UBound(dianxue)
dxname=split(dianxue(x),"|")
if dxname(0)=info(0) then
if DateDiff("s",dxname(1),sj2)<60 then
Response.Write "<script Language=Javascript>alert('�A�w�Q�����I�F�ץީγQ�T���ΩҶˡA���వ����ưڡI�ٳ�" & 60-DateDiff("s",dxname(1),sj2) & "��۰ʸѶ}!');</script>"
response.end
else
dxcz=dianxue(x)&";"
application("ljjh_dianxuename")=replace(application("ljjh_dianxuename"),dxcz,"")
Response.Write "<script Language=Javascript>alert('�ɶ���F�A�A���ץަ۰ʸѶ}�F!');</script>"
response.end
end if
end if
next
end if
Application.UnLock
towho=Trim(Request.Form("towho"))
says=Trim(Request.Form("sy"))
says=replace(says," ","")
if towho="" then towho="�j�a"
if len(towho)>10 then towho=Left(towho,10)
gw=left(towho,1)
if IsNumeric(gw)=true then
zz=split(gw,"��")
gw=1
else 
gw=0
end if  
if Not(gw=1 or towho="�j�a" or towho="�W�ťX��]" or towho="�ԯ�" or towho="�g���j�s" or Instr(useronlinename," "&towho&" ")<>0 or Instr(useronlinename," "&nickmane&" ")<>0) then
Response.Write "<script Language=Javascript>alert('��" & towho & "�����b��ѫǤ��A������o���C');parent.f2.document.af.towho.options[0].value='�j�a';parent.f2.document.af.towho.options[0].text='�j�a';parent.f3.location.reload();</script>"
Response.end
end if
act=0
towhoway=Request.Form("towhoway")
if towhoway<>1 then towhoway=0
addwordcolor=Request.Form("addwordcolor")
sayscolor=Request.Form("sayscolor")
addsays=Request.Form("addsays")
towhoid=Trim(Request.Form("towhoid"))
titleline=Request.Form("titleline")
if InStr(addsays,"for(")<>0 or InStr(addsays,"alert(")<>0  then 
Response.Write "<script language=JavaScript>{alert('��J���~�I');}</script>"
Response.End 
end if

if info(3)<20 or titleline<>1 then titleline=0
if titleline<>1 then titleline=0
if Instr(says,".")<>0 or Instr(says,"/")<>0 or Instr(says,"��")<>0 or Instr(says,"�O")<>0 then titleline=0
if towho="�j�a" or titleline=1 then towhoway=0
if towho="�����`��" then towhoway=1
if len(says)>150 then says=Left(says,150)
if (says="" or says="//") then Response.End
says=Replace(says,"&amp;","&")
//says=lcase(says)
says=Replace(says,"&amp;","&")
//says=lcase(says)
if info(2)<9 then
says=replace(says,"&#","")
says=replace(says,"..","")
says=replace(says,"'","\'")
says=replace(says,chr(34),"\"&chr(34))
says=Replace(says,"<"," ")
says=Replace(says,">"," ")
says=Replace(says,"\x3c"," ")
says=Replace(says,"\x3e"," ")
says=Replace(says,"\074"," ")
says=Replace(says,"\74"," ")
says=Replace(says,"\75"," ")
says=Replace(says,"\76"," ")
says=Replace(says,"&lt"," ")
says=Replace(says,"&gt"," ")
says=Replace(says,"\076"," ")
says=Replace(LCase(says),LCase("con"),"f2uck")
says=Replace(LCase(says),LCase("or"),"����")
says=Replace(LCase(says),LCase("java"),"f2uck")
says=Replace(LCase(says),LCase("www"),"f2uck")
says=Replace(LCase(says),LCase("com"),"f2uck")
says=Replace(LCase(says),LCase("net"),"f2uck")
says=Replace(says,"game","f2uck")
says=Replace(says,"��","f2uck")
says=Replace(says,"asp","f2uck")
says=Replace(says,"d27","f2uck")
says=Replace(says,"66c","f2uck")
says=Replace(says,"����","f2uck")
says=Replace(says,"������","f2uck")
says=Replace(says,"009","f2uck")
says=Replace(says,"����c","f2uck")
says=Replace(says,"jh","f2uck")
says=Replace(says,".to","f2uck")
says=Replace(LCase(says),LCase("http"),"f2uck")
maren=instr(says,"f2uck")
if maren<>0 then
Response.Write "<script Language=Javascript>top.location.href='autokick.asp';</script>"
Response.end
end if
end if
if info(2)<10 then
if instr(says,"�o�P")=0 then 
if InStr(says,"�[��")<>0 then Response.End
end if
end if
'�O�_�԰����Ťj��5�~�i�H�K��
if info(3)>2 and Instr(says," ")=0 and Instr(says,"[img]")<>0 then
if Instr(says,"[img]")<>0 then
if instr(says,"width")<>0 or instr(says,"height")<>0 then Response.End	
gsz=0
Do While Instr(says,"[img]")<>0
says=Replace(says,"[img]","<img src=pic/",1,1)
gsz=gsz+1
Loop
gsy=0
Do While Instr(says,"[/img]")<>0
says=Replace(says,"[/img]",">",1,1)
gsy=gsy+1
Loop
if gsz>gsy then says=says&">"
end if
end if
'�����[�t�]�mxuni�b//��/�ɨ����������
'xuni=0
zj="<a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('[" & info(9) & "]');  target=f2><font color="& addwordcolor & ">" & info(0) & "</font></a>"
br="<a href=javascript:parent.sw('[" & towho & "]'); parent.sws('[" & towhoid & "]'); target=f2><font color=FFD7F4>" & towho & "</font></a>"
if Left(says,2)="//" then
'xuni=1
says=Replace(says,"##",zj,1,3,1)
says=Replace(says,"%%",br,1,3,1)
says="<font color=" & sayscolor & ">" & zj & Trim(mid(says,3,len(says)-2)) & "</font>"
act=1
titleline=0
end if

if left(says,1)="/" then
titleline=0
act=1
towhoway=0
neili=0
i=instr(says,"$")
if i=0 then i=len(says)+1
if trim(replace(mid(says,i+1)," ",""))="" and instr(";�ϥ�;��q;�߸�;�@��;���y;�n��;���H���;�M��;�ץ�;�I��;����;�o��;�Ǳ�;�e��;�U��;�g��;ĵ�i;�[�J;�D�D;�߰�;��j;�s��;����;���;��H;�ذe;",right(left(says,i-1),len(left(says,i-1))-1))<>0 then
Response.Write "<script language=javascript>alert('�i"&right(left(says,i-1),len(left(says,i-1))-1)&"�j�R�O�ާ@���ର�šI');</script>"
Response.End
end if

fnn1=mid(says,i+1)
%><!--#include file="ljfunc/func.asp"--><%
select case left(says,i-1)
case "/�D�D","/title"
%><!--#include file="ljfunc/1.asp"--><%
says="<font color=" & sayscolor & ">"+titl(mid(says,i+1))+"</font>"
towho="�j�a"
towhoid=0
case "/����|���O"
%><!--#include file="ljfunc/give1.asp"--><%
says="<font color=red>�i����|�O�j<font color=" & sayscolor & ">"+give1()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�������"
%><!--#include file="ljfunc/give2.asp"--><%
says="<font color=red>�i��������j<font color=" & sayscolor & ">"+give2()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�����O"
%><!--#include file="ljfunc/give3.asp"--><%
says="<font color=red>�i�����O�j<font color=" & sayscolor & ">"+give3()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/������O"
%><!--#include file="ljfunc/give4.asp"--><%
says="<font color=red>�i������O�j<font color=" & sayscolor & ">"+give4()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/����Z�\"
%><!--#include file="ljfunc/give5.asp"--><%
says="<font color=red>�i����Z�\�j<font color=" & sayscolor & ">"+give5()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�������"
%><!--#include file="ljfunc/give6.asp"--><%
says="<font color=red>�i��������j<font color=" & sayscolor & ">"+give6()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/������s"
%><!--#include file="ljfunc/give7.asp"--><%
says="<font color=red>�i������s�j<font color=" & sayscolor & ">"+give7()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/����y�O"
%><!--#include file="ljfunc/give8.asp"--><%
says="<font color=red>�i����y�O�j<font color=" & sayscolor & ">"+give8()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/����D�w"
%><!--#include file="ljfunc/give9.asp"--><%
says="<font color=red>�i����D�w�j<font color=" & sayscolor & ">"+give9()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/����k�O"
%><!--#include file="ljfunc/give10.asp"--><%
says="<font color=red>�i����k�O�j<font color=" & sayscolor & ">"+give10()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/������"
%><!--#include file="ljfunc/give01.asp"--><%
says="<font color=red>�i������j<font color=" & sayscolor & ">"+give01()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR�|�O"
%><!--#include file="ljfunc/kzd.asp"--><%
says="<font color=red>�i�ʶR�|�O�j<font color=" & sayscolor & ">"+kzd()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR����"
%><!--#include file="ljfunc/kzda.asp"--><%
says="<font color=red>�i�ʶR�����j<font color=" & sayscolor & ">"+kzda()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR�l�u"
%><!--#include file="ljfunc/kzdb.asp"--><%
says="<font color=red>�i�ʶR�l�u�j<font color=" & sayscolor & ">"+kzdb()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR���"
%><!--#include file="ljfunc/kzdc.asp"--><%
says="<font color=red>�i�ʶR���j<font color=" & sayscolor & ">"+kzdc()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR�k�O"
%><!--#include file="ljfunc/kzdd.asp"--><%
says="<font color=red>�i�ʶR�k�O�j<font color=" & sayscolor & ">"+kzdd()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR��O"
%><!--#include file="ljfunc/kzde.asp"--><%
says="<font color=red>�i�ʶR��O�j<font color=" & sayscolor & ">"+kzde()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR�Z�\"
%><!--#include file="ljfunc/kzdf.asp"--><%
says="<font color=red>�i�ʶR�Z�\�j<font color=" & sayscolor & ">"+kzdf()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR����"
%><!--#include file="ljfunc/kzdg.asp"--><%
says="<font color=red>�i�ʶR�����j<font color=" & sayscolor & ">"+kzdg()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR���s"
%><!--#include file="ljfunc/kzdh.asp"--><%
says="<font color=red>�i�ʶR���s�j<font color=" & sayscolor & ">"+kzdh()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR�y�O"
%><!--#include file="ljfunc/kzdi.asp"--><%
says="<font color=red>�i�ʶR�y�O�j<font color=" & sayscolor & ">"+kzdi()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR�D�w"
%><!--#include file="ljfunc/kzdj.asp"--><%
says="<font color=red>�i�ʶR�D�w�j<font color=" & sayscolor & ">"+kzdj()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR2�ŷ|��"
%><!--#include file="ljfunc/kzdk.asp"--><%
says="<font color=red>�i�ʶR2�ŷ|���j<font color=" & sayscolor & ">"+kzdk()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR3�ŷ|��"
%><!--#include file="ljfunc/kzdl.asp"--><%
says="<font color=red>�i�ʶR3�ŷ|���j<font color=" & sayscolor & ">"+kzdl()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR4�ŷ|��"
%><!--#include file="ljfunc/kzdm.asp"--><%
says="<font color=red>�i�ʶR4�ŷ|���j<font color=" & sayscolor & ">"+kzdm()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR6�ũx��"
%><!--#include file="ljfunc/kzdo.asp"--><%
says="<font color=red>�i�ʶR6�ũx���j<font color=" & sayscolor & ">"+kzdo()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR7�ũx��"
%><!--#include file="ljfunc/kzdp.asp"--><%
says="<font color=red>�i�ʶR7�ũx���j<font color=" & sayscolor & ">"+kzdp()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR8�ũx��"
%><!--#include file="ljfunc/kzdq.asp"--><%
says="<font color=red>�i�ʶR8�ũx���j<font color=" & sayscolor & ">"+kzdq()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR9�ũx��"
%><!--#include file="ljfunc/kzdr.asp"--><%
says="<font color=red>�i�ʶR9�ũx���j<font color=" & sayscolor & ">"+kzdr()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR10�ũx��"
%><!--#include file="ljfunc/kzds.asp"--><%
says="<font color=red>�i�ʶR10�ũx���j<font color=" & sayscolor & ">"+kzds()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR���s�Ԥh"
%><!--#include file="ljfunc/kzdt.asp"--><%
says="<font color=red>�i�ʶR���s�Ԥh�j<font color=" & sayscolor & ">"+kzdt()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR����i�h"
%><!--#include file="ljfunc/kzdu.asp"--><%
says="<font color=red>�i�ʶR����i�h�j<font color=" & sayscolor & ">"+kzdu()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�ʶR���Ѯv"
%><!--#include file="ljfunc/kzdv.asp"--><%
says="<font color=red>�i�ʶR���Ѯv�j<font color=" & sayscolor & ">"+kzdv()+"</font>"
towhoway=0
towho=info(0)
towhoid=info(9)
case "/�I��"
%><!--#include file="ljfunc/2.asp"--><%
says="<font color=00AFEF>�i�I�ޡj</font><font color=" & sayscolor & ">"+dian(towhoid,mid(says,i+1),towho)+"</font>"
case "/�ѥ�"
%><!--#include file="ljfunc/3.asp"--><%
says="<font color=00AFEF>�i�ѥޡj</font><font color=" & sayscolor & ">"+jie(mid(says,i+1))+"</font>"
case "/�e��"
%><!--#include file="ljfunc/4.asp"--><%
says="<font color=00AFEF>�i�e���j</font><font color=" & sayscolor & ">"+daipu(mid(says,i+1),towhoid,towho)+"</font>"
case "/���c"
%><!--#include file="ljfunc/5.asp"--><%
says="<font color=00AFEF>�i���c�j</font><font color=" & sayscolor & ">"+zuolao(mid(says,i+1),towhoid,towho)+"</font>"
case "/�ʸT"
%><!--#include file="ljfunc/6.asp"--><%
says="<font color=00AFEF>�i�ʸT�j</font><font color=" & sayscolor & ">"+jianjin(mid(says,i+1),towhoid,towho)+"</font>"
case "/����"
%><!--#include file="ljfunc/7.asp"--><%
says="<font color=00AFEF>�i�Ѱ��ʸT�j</font><font color=" & sayscolor & ">"+shifang(mid(says,i+1))+"</font>"
case "/ĵ�i","/�x���O"
%><!--#include file="ljfunc/9.asp"--><%
if info(2)<7 and info(6)="�x��" then
	says="<font color=00AFEF>�i�x���O�j<font color=" & sayscolor & ">"+jing(mid(says,i+1),towho)+"</font>"
else
	says="<font color=00AFEF>�iĵ�i�j<font color=" & sayscolor & ">"+jing(mid(says,i+1),towho)+"</font>"
end if
case "/���["
%><!--#include file="ljfunc/14c.asp"--><%
says="<font color=00AFEF>�i���[�j<font color=" & sayscolor & ">"+attackc(mid(says,i+1),towhoid)+"</font>"
case "/�g��"
%><!--#include file="ljfunc/14d.asp"--><%
says="<font color=00AFEF>�i�g���j<font color=" & sayscolor & ">"+attackd(mid(says,i+1),towhoid,towho)+"</font>"
case "/�U�r"
%><!--#include file="ljfunc/10.asp"--><%
says="<font color=00AFEF>�i�U�r�j<font color=" & sayscolor & ">"+xiadu(mid(says,i+1),towhoid,towho)+"</font>"
case "/����"
%><!--#include file="ljfunc/11.asp"--><%
says="<font color=00AFEF>�i�s���j<font color=" & sayscolor & ">"+touqian(towhoid,towho)+"</font>"
case "/��F�o�a"
%><!--#include file="ljfunc/53.asp"--><%
says="<font color=00AFEF>�i����ʡj<font color=" & sayscolor & ">"+laobao(towhoid,towho)+"</font>"
case "/����"
%><!--#include file="ljfunc/533.asp"--><%
says="<font color=00AFEF>�i������ʡj<font color=" & sayscolor & ">"+saohuang(towhoid,towho)+"</font>"
case "/�l�P�j�k"
%><!--#include file="ljfunc/12.asp"--><%
says="<font color=00AFEF>�i�l�P�j�k�j<font color=" & sayscolor & ">"+xxdf(towhoid,towho)+"</font>"
case "/���Y"
%><!--#include file="ljfunc/13.asp"--><%
says="<font color=00AFEF>�i�o�g�t���j<font color=" & sayscolor & ">"+touzi(mid(says,i+1),towhoid,towho)+"</font>"
case "/�o��"
%><!--#include file="ljfunc/14.asp"--><%
says="<font color=00AFEF>�i"&info(5)&"�Z�ǡj<font color=" & sayscolor & ">"+attack(mid(says,i+1),towhoid,towho)+"</font>"
case "/�m��"
%><!--#include file="ljfunc/14a.asp"--><%
says="<font color=00AFEF>�i�m���_���j<font color=" & sayscolor & ">"+qiang(mid(says,i+1),towhoid,towho)+"</font>"
case "/�ǰe�k�O"
%><!--#include file="ljfunc/45002.asp"--><%
says="<font color=00AFEF>�i�ǰe�k�O�j<font color=" & sayscolor & ">"+songfali(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/�M�����"
%><!--#include file="ljfunc/45003.asp"--><%
says="<font color=green>[�M�����]<font color=" & sayscolor & ">"+xunshuijing(mid(says,i+1),towhoid,towho)+"</font>"
case "/�M��k��"
%><!--#include file="ljfunc/45005.asp"--><%
says="<font color=green>[�M��k��]<font color=" & sayscolor & ">"+xunfaqi(mid(says,i+1),towhoid,towho)+"</font>"
case "/�M��y�P"
%><!--#include file="ljfunc/4000.asp"--><%
says="<font color=green>[�M��y�P]<font color=" & sayscolor & ">"+bbbbbb(mid(says,i+1),towhoid,towho)+"</font>"
case "/�M��y�P�\"
%><!--#include file="ljfunc/4001.asp"--><%
says="<font color=green>[�M��y�P�\]<font color=" & sayscolor & ">"+cccccc(mid(says,i+1),towhoid,towho)+"</font>"
case "/�M���s�]"
%><!--#include file="ljfunc/4002.asp"--><%
says="<font color=green>[�M���s�]]<font color=" & sayscolor & ">"+dddddd(mid(says,i+1),towhoid,towho)+"</font>"
case "/�M��|�O"
%><!--#include file="ljfunc/9000.asp"--><%
says="<font color=green>[�M��|�O]<font color=" & sayscolor & ">"+dog(mid(says,i+1),towhoid,towho)+"</font>"
case "/�M�����"
%><!--#include file="ljfunc/9001.asp"--><%
says="<font color=green>[�M�����]<font color=" & sayscolor & ">"+dogdog(mid(says,i+1),towhoid,towho)+"</font>"
case "/�]�O�p��"
%><!--#include file="ljfunc/45006.asp"--><%
says="<font color=00AFEF>�i�]�O�p�ۡj<font color=" & sayscolor & ">"+molishi(mid(says,i+1),towhoid)+"</font>"
case "/�ͤ�J�|"
%><!--#include file="ljfunc/450018.asp"--><%
says="<font color=00AFEF>�i�ͤ�J�|�j<font color=" & sayscolor & ">"+shendangao(mid(says,i+1),towhoid)+"</font>"
case "/�t���_��"
%><!--#include file="ljfunc/45015.asp"--><%
says="<font color=00AFEF>�i�t���_�ۡj<font color=" & sayscolor & ">"+peibashi(mid(says,i+1),towhoid)+"</font>"
case "/�o�g�l�u"
%><!--#include file="ljfunc/450017.asp"--><%
says="<font color=00AFEF>�i�o�g�l�u�j<font color=" & sayscolor & ">"+fashezi(mid(says,i+1),towhoid,towho)+"</font>"
case "/�Ǳ�"
%><!--#include file="ljfunc/15.asp"--><%
says="<font color=00AFEF>�i�Ǳ¡j<font color=" & sayscolor & ">"+cuan(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/���ܨp��"
%><!--#include file="ljfunc/888.asp"--><%
says="<font color=00AFEF>�i���ܨp��j<font color=" & sayscolor & ">"+tracksl()+"</font>" 
towhoway=1
towho=info(0)
towhoid=info(9)
case "/��������" 
%><!--#include file="ljfunc/999.asp"--><%
says="<font color=00AFEF>�i�������ܡj<font color=" & sayscolor & ">"+clearsl()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�T���}��"
%><!--#include file="ljfunc/666.asp"--><%
says="<font color=00AFEF>�i�T���}���j<font color=" & sayscolor & ">"+jingda()+"</font>" 
case "/����"
%><!--#include file="ljfunc/ying.asp"--><%
says="<font color=00AFEF>�i�����}���j<font color=" & sayscolor & ">"+yingsheng()+"</font>" 
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�e��"
%><!--#include file="ljfunc/16.asp"--><%
says="<font color=00AFEF>�i�e���j<font color=" & sayscolor & ">"+give(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/�e�|�O"
%><!--#include file="ljfunc/2016.asp"--><%
says="<font color=00AFEF>�i�ذe�|�O�j<font color=" & sayscolor & ">"+eeee(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/�e����"
%><!--#include file="ljfunc/2017.asp"--><%
says="<font color=00AFEF>�i�ذe�����j<font color=" & sayscolor & ">"+ffff(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/���W�B�m"
%><!--#include file="ljfunc/161.asp"--><%
says="<font color=00AFEF>�i���W�B�m�j<font color=" & sayscolor & ">"+mengui(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/�@��"
%><!--#include file="ljfunc/611.asp"--><%
says="<font color=00AFEF>�i����@�ڡj</font><font color=" & sayscolor & ">"+fakuan(mid(says,i+1),towhoid,towho)+"</font>"
case "/�j��X���N"
%><!--#include file="ljfunc/2000.asp"--><%
says="<font color=00AFEF>�i�j��X���N�j</font><font color=" & sayscolor & ">"+aa(mid(says,i+1),towhoid,towho)+"</font>"
case "/�a����޳N"
%><!--#include file="ljfunc/2001.asp"--><%
says="<font color=00AFEF>�i�a����޳N�j</font><font color=" & sayscolor & ">"+bb(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����[�޳N"
%><!--#include file="ljfunc/2002.asp"--><%
says="<font color=00AFEF>�i�����[�޳N�j</font><font color=" & sayscolor & ">"+cc(mid(says,i+1),towhoid,towho)+"</font>"
case "/�·t���N"
%><!--#include file="ljfunc/2003.asp"--><%
says="<font color=00AFEF>�i�·t���N�j</font><font color=" & sayscolor & ">"+dd(mid(says,i+1),towhoid,towho)+"</font>"
case "/�·t��N"
%><!--#include file="ljfunc/2004.asp"--><%
says="<font color=00AFEF>�i�·t��N�j</font><font color=" & sayscolor & ">"+ee(mid(says,i+1),towhoid,towho)+"</font>"
case "/�·t����N"
%><!--#include file="ljfunc/2005.asp"--><%
says="<font color=00AFEF>�i�·t����N�j</font><font color=" & sayscolor & ">"+ff(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����[���N"
%><!--#include file="ljfunc/2006.asp"--><%
says="<font color=00AFEF>�i�����[���N�j</font><font color=" & sayscolor & ">"+aaa(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����[���N"
%><!--#include file="ljfunc/2007.asp"--><%
says="<font color=00AFEF>�i�����[���N�j</font><font color=" & sayscolor & ">"+bbb(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����[��N"
%><!--#include file="ljfunc/2008.asp"--><%
says="<font color=00AFEF>�i�����[��N�j</font><font color=" & sayscolor & ">"+ccc(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����[���N"
%><!--#include file="ljfunc/5001.asp"--><%
says="<font color=00AFEF>�i�����[���N�j</font><font color=" & sayscolor & ">"+abcabc(mid(says,i+1),towhoid,towho)+"</font>"
case "/�·t����N"
%><!--#include file="ljfunc/2009.asp"--><%
says="<font color=00AFEF>�i�·t����N�j</font><font color=" & sayscolor & ">"+ddd(mid(says,i+1),towhoid,towho)+"</font>"
case "/�·t��|�N"
%><!--#include file="ljfunc/2010.asp"--><%
says="<font color=00AFEF>�i�·t��|�N�j</font><font color=" & sayscolor & ">"+eee(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����[���N"
%><!--#include file="ljfunc/2011.asp"--><%
says="<font color=00AFEF>�i�����[���N�j</font><font color=" & sayscolor & ">"+fff(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����[�|�N"
%><!--#include file="ljfunc/2012.asp"--><%
says="<font color=00AFEF>�i�����[�|�N�j</font><font color=" & sayscolor & ">"+aaaa(mid(says,i+1),towhoid,towho)+"</font>"
case "/�հ��ʯv�N"
%><!--#include file="ljfunc/2013.asp"--><%
says="<font color=00AFEF>�i�հ��ʯv�N�j</font><font color=" & sayscolor & ">"+bbbb(mid(says,i+1),towhoid,towho)+"</font>"
case "/�I�}�ѯv�N"
%><!--#include file="ljfunc/2014.asp"--><%
says="<font color=00AFEF>�i�I�}�ѯv�N�j</font><font color=" & sayscolor & ">"+cccc(mid(says,i+1),towhoid,towho)+"</font>"
case "/�a������N"
%><!--#include file="ljfunc/5002.asp"--><%
says="<font color=00AFEF>�i�a������N�j</font><font color=" & sayscolor & ">"+gogo(mid(says,i+1),towhoid,towho)+"</font>"
case "/�٭��N"
%><!--#include file="ljfunc/2015.asp"--><%
says="<font color=00AFEF>�i�٭��N�j</font><font color=" & sayscolor & ">"+dddd(mid(says,i+1),towhoid,towho)+"</font>"
case "/���W�B�m"
%><!--#include file="ljfunc/161.asp"--><%
says="<font color=00AFEF>�i���W�B�m�j<font color=" & sayscolor & ">"+mengui(mid(says,i+1)+0,towhoid,towho)+"</font>"
case "/�@��"
%><!--#include file="ljfunc/611.asp"--><%
says="<font color=00AFEF>�i����@�ڡj</font><font color=" & sayscolor & ">"+fakuan(mid(says,i+1),towhoid,towho)+"</font>"
case "/�ʦD"
%><!--#include file="ljfunc/612.asp"--><%
says="<font color=00AFEF>�i����ʦD�j</font><font color=" & sayscolor & ">"+dongxin(mid(says,i+1),towhoid,towho)+"</font>"
case "/�ذe"
%><!--#include file="ljfunc/17.asp"--><%
says="<font color=00AFEF>�i�ذe�j<font color=" & sayscolor & ">"+zen(mid(says,i+1),towhoid,towho)+"</font>"
case "/���"
%><!--#include file="ljfunc/20031.asp"--><%
says="<font color=00AFEF>�i��󪫫~�j<font color=" & sayscolor & ">"+diu(mid(says,i+1),towhoid,towho)+"</font>"
case "/�[�J"
%><!--#include file="ljfunc/19.asp"--><%
says="<font color=00AFEF>�i�[�J�j<font color=" & sayscolor & ">"+join(mid(says,i+1))+"</font>"
case "/�U��"
%><!--#include file="ljfunc/20.asp"--><%
says="<font color=00AFEF>�i�U�ʡj<font color=" & sayscolor & ">"+cefen(mid(says,i+1),towhoid,towho)+"</font>"
case "/���i"
%><!--#include file="ljfunc/21.asp"--><%
says=gong(mid(says,i+1))
case "/��H"
%><!--#include file="ljfunc/22.asp"--><%
says="<font color=00AFEF>�i��H�j<font color=" & sayscolor & ">"+tiren(towhoid,towho,mid(says,i+1))+"</font>"
case "/����ip"
%><!--#include file="ljfunc/22001.asp"--><%
says="<font color=00AFEF>�i����ip�j<font color=" & sayscolor & ">"+jiafeng(towhoid,towho,mid(says,i+1))+"</font>"
case "/����ip"
%><!--#include file="ljfunc/22002.asp"--><%
says="<font color=00AFEF>�i����ip�j</font><font color=" & sayscolor & ">"+jiansu(mid(says,i+1))+"</font>"
case "/�߰�"
%><!--#include file="ljfunc/23.asp"--><%
says=xindong(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/��j"
%><!--#include file="ljfunc/7010.asp"--><%
says=rock(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/�ť��r"
%><!--#include file="ljfunc/24.asp"--><%
says=fangda(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/�����r"
%><!--#include file="ljfunc/1000.asp"--><%
says=a(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/����r"
%><!--#include file="ljfunc/1001.asp"--><%
says=b(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/�����r"
%><!--#include file="ljfunc/1002.asp"--><%
says=c(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/�����r"
%><!--#include file="ljfunc/1003.asp"--><%
says=d(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/�����r"
%><!--#include file="ljfunc/1004.asp"--><%
says=e(mid(says,i+1))
towho="�j�a"
towhoid=0
case "/���v"
%><!--#include file="ljfunc/25.asp"--><%
says="<font color=00AFEF>�i���v�j<font color=" & sayscolor & ">"+bais(towhoid,towho)+"</font>"
case "/�U�s"
%><!--#include file="ljfunc/251.asp"--><%
says="<font color=00AFEF>�i�U�s�j<font color=" & sayscolor & ">"+xiashan(towhoid,towho)+"</font>"
case "/���{"
%><!--#include file="ljfunc/26.asp"--><%
says="<font color=00AFEF>�i���{�j<font color=" & sayscolor & ">"+stu(towho)+"</font>"
case "/�D�c","/�D�p�Ѥ�"
%><!--#include file="ljfunc/00.asp"--><%
says="<font color=00AFEF>�i�D�c�j<font color=" & sayscolor & ">"+qqq(towhoid,towho)+"</font>"
case "/�P�N"
%><!--#include file="ljfunc/000.asp"--><%
says="<font color=00AFEF>�i�P�N�j<font color=" & sayscolor & ">"+dy(towho)+"</font>"
case "/��D����"
%><!--#include file="ljfunc/311.asp"--><%
says="<font color=00AFEF>�i��D�����j</font><font color=" & sayscolor & ">"+quxiaoa()+"</font>"
case "/����"
%><!--#include file="ljfunc/27.asp"--><%
says="<font color=red>�i�����j<font color=" & sayscolor & ">"+dazhuo()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/����"
%><!--#include file="ljfunc/28.asp"--><%
says="<font color=red>�i���ءj<font color=" & sayscolor & ">"+bimu()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�ײߪZ�\"
%><!--#include file="ljfunc/71.asp"--><%
says="<font color=00AFEF>�i�ײߪZ�\�j</font><font color=" & sayscolor & ">"+wuxiu(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/���I"
%><!--#include file="ljfunc/canchan.asp"--><%
says="<font color=red>�i�s�}���I�j<font color=" & sayscolor & ">"+canchan()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/����"
%><!--#include file="ljfunc/shenci.asp"--><%
says="<font color=red>�i�F�C����j<font color=" & sayscolor & ">"+shenci()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�׾i"
%><!--#include file="ljfunc/xiuyang.asp"--><%
says="<font color=red>�i�פ߾i�ʡj<font color=" & sayscolor & ">"+xiuyang()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�g��"
%><!--#include file="ljfunc/29.asp"--><%
says="<font color=red>�i�ǰe�g��j<font color=" & sayscolor & ">"+jingyan(mid(says,i+1),towhoid,towho)+"</font>"
case "/�m�Z"
%><!--#include file="ljfunc/30.asp"--><%
says="<font color=red>�i�m�Z�j<font color=" & sayscolor & ">"+lianwu()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/���R"
%><!--#include file="ljfunc/59.asp"--><%
says="<font color=00AFEF>�i�P�k��ɡj<font color=" & sayscolor & ">"+pingmin(mid(says,i+1),towhoid,towho)+"</font>"
case "/����"
%><!--#include file="ljfunc/61.asp"--><%
says="<font color=red>�i����^�K�j<font color=" & sayscolor & ">"+yiliao()+"</font>"
case "/���e"
%><!--#include file="ljfunc/62.asp"--><%
says="<font color=red>�i���e�j�k�j<font color=" & sayscolor & ">"+meirong()+"</font>"
case "/�ЪZ"
%><!--#include file="ljfunc/63.asp"--><%
says="<font color=red>�i�оɪZ�\�j<font color=" & sayscolor & ">"+junguan()+"</font>"
case "/�ɮ�"
%><!--#include file="ljfunc/64.asp"--><%
says="<font color=red>�i�C���F��j<font color=" & sayscolor & ">"+qigongshi()+"</font>"

case "/�s��"
%><!--#include file="ljfunc/31.asp"--><%
says="<font color=00AFEF>�i�s���j<font color=" & sayscolor & ">"+putyin(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/����"
%><!--#include file="ljfunc/32.asp"--><%
says="<font color=00AFEF>�i�����j<font color=" & sayscolor & ">"+getyin(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�s�k�O"
%><!--#include file="ljfunc/450020.asp"--><%
says="<font color=00AFEF>�i�s�k�O�j<font color=" & sayscolor & ">"+putfali(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/���k�O"
%><!--#include file="ljfunc/450021.asp"--><%
says="<font color=00AFEF>�i���k�O�j<font color=" & sayscolor & ">"+getfali(mid(says,i+1))+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/���"
%><!--#include file="ljfunc/33.asp"--><%
says="<font color=00AFEF>�i���j<font color=" & sayscolor & ">"+zzyin(mid(says,i+1),towhoid,towho)+"</font>"
case "/�׽m"
%><!--#include file="ljfunc/35.asp"--><%
says="<font color=red>�i�_���׽m�j<font color=" & sayscolor & ">"+xiulian()+"</font>"
case "/�ɨ�"
%><!--#include file="ljfunc/36.asp"--><%
says="<font color=red>�i�ɫ¤O���j<font color=" & sayscolor & ">"+baodou()+"</font>"
case "/�ץ�"
%><!--#include file="ljfunc/37.asp"--><%
says="<font color=00AFEF>�i�ץޡj<font color=" & sayscolor & ">"+ya(towhoid,towho,mid(says,i+1))+"</font>"
case "/�٭�"
%><!--#include file="ljfunc/8.asp"--><%
says="<font color=00AFEF>�i�٭��j<font color=" & sayscolor & ">"+zhanshou(mid(says,i+1),towhoid)+"</font>"
case "/��q"
%><!--#include file="ljfunc/39.asp"--><%
says=nuhou(mid(says,i+1))
towho="�j�a"
case "/�߸�"
%><!--#include file="ljfunc/40.asp"--><%
says=xintiao(mid(says,i+1))
towho="�j�a"
case "/�dip"
%><!--#include file="ljfunc/41.asp"--><%
says=seeip(towhoid,towho)
towhoway=1
towho=info(0)
case "/���y"
%><!--#include file="ljfunc/42.asp"--><%
says="<font color=00AFEF>�i���y�j<font color=" & sayscolor & ">"+jiangli(mid(says,i+1),towhoid,towho)+"</font>"
case "/�x�����y"
%><!--#include file="ljfunc/jianglila.asp"--><%
says="<font color=00AFEF>�i�x�����y�j<font color=" & sayscolor & ">"+jianglila(mid(says,i+1),towhoid,towho)+"</font>"
case "/���"
%><!--#include file="ljfunc/43.asp"--><%
says="<font color=red>�i����d�ݡj<font color=" & sayscolor & ">"+seejj()+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�ϥ�"
%><!--#include file="ljfunc/44.asp"--><%
says="<font color=00AFEF>�i�d���j<font color=" & sayscolor & ">"+kapian(mid(says,i+1),towhoid,towho)+"</font>"
case "/�p��"
%><!--#include file="ljfunc/45.asp"--><%
says="<font color=00AFEF>�i�p���j<font color=" & sayscolor & ">"+tdx(towhoid,towho)+"</font>"
case "/��p��"
%><!--#include file="ljfunc/46.asp"--><%
says="<font color=00AFEF>�i��p���j<font color=" & sayscolor & ">"+pk(towhoid,towho)+"</font>"
case "/�Ĥl����"
%><!--#include file="ljfunc/51.asp"--><%
says="<font color=00AFEF>�i�Ĥl���ѡj<font color=" & sayscolor & ">"+hza(towhoid,towho)+"</font>"
case "/�аV�Ĥl"
%><!--#include file="ljfunc/52.asp"--><%
says="<font color=00AFEF>�i�аV�Ĥl�j<font color=" & sayscolor & ">"+hzb(towhoid,towho)+"</font>"
case "/���H���"
%><!--#include file="ljfunc/54.asp"--><%
says="<font color=00AFEF>�i���H��աj<font color=" & sayscolor & ">"+grdb(mid(says,i+1),towhoid,towho)+"</font>"
case "/�D�B"
%><!--#include file="ljfunc/47.asp"--><%
says="<font color=00AFEF>�i�D�B�j<font color=" & sayscolor & ">"+qiuhun(mid(says,i+1),towhoid,towho)+"</font>"
case "/�P��"
%><!--#include file="ljfunc/jiemen.asp"--><%
says="<font color=00AFEF>�i���������j<font color=" & sayscolor & ">"+jiemen(towhoid,towho)+"</font>"
case "/�Ѱ��P��"
%><!--#include file="ljfunc/jiecmen.asp"--><%
says="<font color=00AFEF>�i�Ѱ��P���j<font color=" & sayscolor & ">"+jiecmen(mid(says,i+1),towhoid,towho)+"</font>"
case "/�����j��"
%><!--#include file="ljfunc/bangpai.asp"--><%
says="<font color=00AFEF>�i�����j�ԡj<font color=" & sayscolor & ">"+bangpai(towhoid,towho)+"</font>"
case "/�����J"
%><!--#include file="duju/dpk-ask.asp"--><%
says="<font color=00AFEF>�i�����J�j<font color=" & sayscolor & ">"+dpkask(mid(says,i+1),towhoid,towho)+"</font>"
case "/�o�P"
%><!--#include file="duju/dpkfp.asp"--><%
says="<font color=00AFEF>�i�o�P�j<font color=" & sayscolor & ">"+dpkfp(mid(says,i+1),towhoid,towho)+"</font>"
case "/���±N"
%><!--#include file="duju/dmj-ask.asp"--><%
says="<font color=00AFEF>�i���±N�j<font color=" & sayscolor & ">"+dmjask(mid(says,i+1),towhoid,towho)+"</font>"
case "/�X�P"
%><!--#include file="duju/dmjfp.asp"--><%
says="<font color=00AFEF>�i�X�P�j<font color=" & sayscolor & ">"+dmjfp(mid(says,i+1),towhoid,towho)+"</font>"
case "/�N�P"
%><!--#include file="duju/dmjmp.asp"--><%
says="<font color=00AFEF>�i�N�P�j<font color=" & sayscolor & ">"+dmjGetmj(towhoid,towho)+"</font>"
case "/��P"
%><!--#include file="duju/dmjcp.asp"--><%
says="<font color=00AFEF>�i��P�j<font color=" & sayscolor & ">"+dmjPeng(mid(says,i+1),towhoid,towho)+"</font>"
case "/��P"
%><!--#include file="duju/dmjcp.asp"--><%
says="<font color=00AFEF>�i��P�j<font color=" & sayscolor & ">"+dmjPeng(mid(says,i+1),towhoid,towho)+"</font>"
case "/�ݵP"
%><!--#include file="duju/dmjwp.asp"--><%
says="<font color=00AFEF>�i�ݵP�j<font color=" & sayscolor & ">"+dmjwp(towho)+"</font>"
case "/��D"
%><!--#include file="ljfunc/471.asp"--><%
says="<font color=red>�i�ͦ��M���j<font color=" & sayscolor & ">"+diantiao(mid(says,i+1),towhoid,towho)+"</font>"
case "/�ۦ�"
%><!--#include file="ljfunc/4711.asp"--><%
says="<font color=red>�i�s�̤ۧl�j<font color=" & sayscolor & ">"+jiamen(mid(says,i+1),towhoid,towho)+"</font>"
case "/���B"
%><!--#include file="ljfunc/48.asp"--><%
says="<font color=00AFEF>�i���B�����j<font color=" & sayscolor & ">"+xingyu(mid(says,i+1))+"</font>"
if instr(says,"����")=0 then
towhoway=1
towho=info(0)
towhoid=info(9)
end if
case "/���A"
%><!--#include file="ljfunc/18.asp"--><%
says="<font color=00AFEF>�i���A�j<font color=" & sayscolor & ">"&info(0)&"�o�즿�򱡳�:"+getstat(towhoid,towho)+"</font>"
towhoway=1
towho=info(0)
towhoid=info(9)
case "/�߱�"
%><!--#include file="ljfunc/55.asp"--><%
says="<font color=00AFEF>�i�߱��j<font color=" & sayscolor & ">"+xinqing(mid(says,i+1))+"</font>"
towho="mingdan"
case "/�������~"
Response.Write "<script language=javascript>window.open('flwp.asp?fl="& mid(says,i+1) &"');</script>"
Response.End
case else
Response.Write "<script Language=Javascript>alert('�z��J���R�O���~�I');</script>"
response.end
end select
nowcz=0
if len(info(0))<=len(towho)  then
if instr(towho,info(0))>0 then nowcz=1
end if
if nowcz=0 then
says=Replace(says,info(0),"<a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('[" & info(9) & "]'); target=f2  ><font  color="&addwordcolor&">"& info(0) &"</font></a>")
end if
end if
addsays=addsays&"��"
Session("ljjh_lasttime")=sj2
says=Replace(says,"\","\\")
says=Replace(says,"/","\/")
says=Replace(says,chr(34),"\"&chr(34))
if titleline=1 then
Application.Lock
Application("ljjh_title"&session("nowinroom"))=Left(says,32) & "<font color=00FFFF style=font-size:11pt>�]" & info(0) & "�A" & sj & "�^</font>"
Application.UnLock
call showsay()
end if
Application.Lock
	sd=Application("ljjh_sd")
	line=int(Application("ljjh_line"))
	Application("ljjh_line")=line+1
Application.UnLock
for i=1 to 190
  sd(i)=sd(i+10)
next
sd(191)=line+1
if act=1 then
sd(192)=-1
else
sd(192)=info(9)
end if
sd(193)=towhoway
sd(194)=info(0)
sd(195)=towho
sd(196)=addwordcolor
sd(197)=sayscolor
sd(198)=addsays
sd(199)=says&"..."
sd(200)=session("nowinroom")
Application.Lock
Application("ljjh_sd")=sd
Application.UnLock
if chatinfo(6)<>0 then
randomize()
rnd1=int(rnd*750)
sayyq=""
if rnd1<=2 then
Set conn=Server.CreateObject("ADODB.CONNECTION")
conn.open Application("ljjh_usermdb")
conn.execute "Update �Τ� set �O�@=false where �|������=0"
conn.close
set conn=nothing
sayyq="<center><font size=4 color=red>�i�@�}�j�����,�Ҧ����D�|�����Q�Ѱ��O�@�F..�Ȧ����ֶ}�O�@��!�j</center><bgsound src=wav/003.MID loop=1></font>"
end if
if rnd1>=21 and rnd1<=25 then
abc="<a href='qiang.asp' target='d'><img src='img/tank.gif' border='0'></a>"
sayyq="<font color=red>�@�����j�X�{�b��ѫǸ̡A�j�a�ַm��!</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
Application.Lock
Application("ljjh_qiang")=1
Application.UnLock
end if
if rnd1>=26 and rnd1<=30 then
Application.Lock
Application("ljjh_jiaofei")="�g��"
Application.UnLock
sayyq="<font color=red>�i�����j</font><font color=red>��M�@�s�g��<img src=img/jiao.gif>���J����m�T�A�а���̧֥h�ϭ�ڡC</font>"
end if
if rnd1>=31 and rnd1<=35 then
jstl=int(rnd*5000)+100
Application.Lock
Application("ljjh_kl")=jstl
Application.UnLock
abc="<a href='kl.asp' target='d'><img src='img/shenmo.GIF' border='0' width='54' height='54'></a>"
sayyq="<font color=red>��M�������}�}  �����ǧr�I���@�k�l�y�s�A�@�Y�������i��ѫǡA������O�G+"&jstl&"</font><br><marquee width=100% behavior=alternate scrollamount=15>"&abc&"</marquee>"
end if
	if sayyq<>"" then
		Application.Lock
		sd=Application("ljjh_sd")
		line=int(Application("ljjh_line"))
		Application("ljjh_line")=line+1
		Application.UnLock
		for i=1 to 190
		sd(i)=sd(i+10)
		next
		sd(191)=line+1
		sd(192)=-1
		sd(193)=0
		sd(194)=info(0)
		sd(195)="�j�a"
		sd(196)=addwordcolor
		sd(197)=sayscolor
		sd(198)=addsays
		sd(199)=sayyq&"..."
		sd(200)=session("nowinroom")
		Application.Lock
		Application("ljjh_sd")=sd
		Application.UnLock
	end if
end if
call showsay()
Sub showsay()
filname=Session("ljjh_filname")
Response.Write "<script Language=JavaScript>"
sd=Application("ljjh_sd")
userline=int(Session("ljjh_line"))
newuserline=0
Dim show()
j=1
for i=1 to 200 step 10
newuserline=sd(i)
if sd(i)>userline and cstr(sd(i+9))=cstr(session("nowinroom")) and (session("slshow")=1 or (sd(i+2)="0" or sd(i+4)="�j�a" or sd(i+2)="1" and (CStr(sd(i+3))=CStr(info(0)) or CStr(sd(i+4))=CStr(info(0)))) and Instr(filname," "&sd(i+3)&",")=0) then
Redim Preserve show(j),show(j+1),show(j+2),show(j+3),show(j+4),show(j+5),show(j+6),show(j+7),show(j+8),show(j+9)
show(j)=sd(i)
show(j+1)=sd(i+1)
show(j+2)=sd(i+2)
show(j+3)=sd(i+3)
show(j+4)=sd(i+4)
show(j+5)=sd(i+5)
show(j+6)=sd(i+6)
show(j+7)=sd(i+7)
show(j+8)=sd(i+8)
show(j+9)=sd(i+9)
j=j+10
end if
next
if j>1  then
		for i=1 to UBound(show) step 10
		Response.Write "parent.sh(" & show(i+1) & ","& towhoid & "," & show(i+2) & "," & chr(34) & show(i+3) & chr(34) & "," & chr(34) & show(i+4) & chr(34) & "," & chr(34) & show(i+5) & chr(34) & "," & chr(34) & show(i+6) & chr(34) & "," & chr(34) & show(i+7) & chr(34) & "," & chr(34) & show(i+8) & chr(34) & ");"
		next
end if
if titleline=1 then Response.Write "parent.t.location.reload();"
Response.Write "</script>"
if newuserline>userline then Session("ljjh_line")=newuserline
Response.End
End Sub
%>