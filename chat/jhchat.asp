<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.Expires=0
Response.ExpiresAbsolute = Now() - 1
Response.AddHeader "Pragma","No-Cache"
Response.AddHeader "Cache-Control","Private"
Response.CacheControl = "No-Cache"
Response.Expires=0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
ljjh_zm= info(4) & "|" & info(5) & "|"&info(0)&"|"&info(1)&"|"&info(12)&"|" & info(6) &"|"&info(3)&"|"&info(9)&"|"&info(10)&"|"&info(8)&"|"&info(14)
scrname=Request.ServerVariables("SCRIPT_NAME")
if Instr(LCase(scrname),"jhchat.asp")=0 then Response.Redirect "../error.asp?id=002"
Session("ljjh_filname")=" ,"
ljjh_roominfo=split(Application("ljjh_room"),";")
chatroomnum=ubound(ljjh_roominfo)-1
allhttp=LCase(Request.ServerVariables("ALL_HTTP"))
'�o�@��O�˴��Τ�O�_�����άO��L��
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
sql="SELECT �ާ@�ɶ�,���A,�t�� FROM �Τ� WHERE ���A='���`' and id="&info(9)
rs.open sql,conn,1,3
if rs.Eof and rs.Bof then
rs.close
set rs=nothing
conn.close
set conn=nothing
Response.Redirect "../error.asp?id=456"
end if
if Session("ljjh_inthechat")<>"1" then
conn.execute "update �Τ� set �ާ@�ɶ�=now() where id="&info(9)
end if
zt=rs("���A")
pe=rs("�t��")
rs.close
set rs=nothing
conn.close
set conn=nothing

if zt="" or zt="��" or zt="�v" or zt="�L" or zt="�c" or zt="��" or zt="�ʸT" then  Response.Redirect "../error.asp?id=456"
do while DateDiff("s",Application("ljjh_czsj"),now())<3
loop
Application.Lock
Application("ljjh_czsj")=now()
Application.UnLock
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
sj=n & "-" & y & "-" & r & " " & s & ":" & f & ":" & m
t=s & ":" & f & ":" & m
Session("ljjh_lastsaytime")=sj
if Session("ljjh_inthechat")<>"1" then
for i=0 to chatroomnum
if InStr(LCase(Application("ljjh_useronlinename"&i))," " & LCase(info(0)) & " ")<>0 then Response.Redirect "../error.asp?id=300"
next
Session("ljjh_inthechat")="1"
Session("ljjh_savetime")=now()
Session("ljjh_lasttime")=sj
Application.Lock
dim newonlinelist()
onlinelist=Application("ljjh_onlinelist"&session("nowinroom"))
useronlinename=""
onliners=0
js=1
yjl=0
ubl=UBound(onlinelist)
for i=1 to ubl step 6
if DateDiff("n",onlinelist(i+5),sj)<=60 then
if yjl=0 and len(onlinelist(i+1))>len(info(0)) then
yjl=1
onliners=onliners+2
useronlinename=useronlinename & " " & info(0) & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5),newonlinelist(js+6),newonlinelist(js+7),newonlinelist(js+8),newonlinelist(js+9),newonlinelist(js+10),newonlinelist(js+11)
newonlinelist(js)=info(9)
newonlinelist(js+1)=info(0)
newonlinelist(js+2)=info(1)&info(4)
newonlinelist(js+3)=ljjh_zm
newonlinelist(js+4)=sj
newonlinelist(js+5)=sj
newonlinelist(js+6)=onlinelist(i)
newonlinelist(js+7)=onlinelist(i+1)
newonlinelist(js+8)=onlinelist(i+2)
newonlinelist(js+9)=onlinelist(i+3)
newonlinelist(js+10)=onlinelist(i+4)
newonlinelist(js+11)=onlinelist(i+5)
js=js+12
else
onliners=onliners+1
useronlinename=useronlinename & " " & onlinelist(i+1)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=onlinelist(i)
newonlinelist(js+1)=onlinelist(i+1)
newonlinelist(js+2)=onlinelist(i+2)
newonlinelist(js+3)=onlinelist(i+3)
newonlinelist(js+4)=onlinelist(i+4)
newonlinelist(js+5)=onlinelist(i+5)
js=js+6
end if
end if
next
if yjl=0 then
onliners=onliners+1
useronlinename=useronlinename & " " & info(0)
Redim Preserve newonlinelist(js),newonlinelist(js+1),newonlinelist(js+2),newonlinelist(js+3),newonlinelist(js+4),newonlinelist(js+5)
newonlinelist(js)=info(9)
newonlinelist(js+1)=info(0)
newonlinelist(js+2)=info(1)&info(4)
newonlinelist(js+3)=ljjh_zm
newonlinelist(js+4)=sj
newonlinelist(js+5)=sj
end if
useronlinename=useronlinename&" "
if onliners=0 then
dim listnull(0)
Application("ljjh_onlinelist"&session("nowinroom"))=listnull
else
Application("ljjh_onlinelist"&session("nowinroom"))=newonlinelist
end if
Application("ljjh_useronlinename"&session("nowinroom"))=useronlinename
Application("ljjh_chatrs"&session("nowinroom"))=onliners
Application.UnLock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Session("ljjh_line")=line

'�����i�J�O�_��ܫH���A�p�G����ܧ令'��������
if Instr(1,Application("ljjh_ying"),"|"& info(0) & "|")=0  then
Application.Lock
Application("ljjh_line")=line+1
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=0
sd(194)=info(0)
sd(195)="mingdan"
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
randomize timer
chatdj=int(rnd*9)+1
aszcc=int(rnd*9)+1
Select Case chatdj
case 1
ljjh_userinto="�b���򪺤j�D�W,%%�M�ۤ@�Y�����P�����ޫ��A�⮳�@�ڬh����A�Ө�F"&Application("ljjh_chatroomname")&",���ħr�B���r����,�榣�����:���U��S���A�p�̦�§!��"
case 2
ljjh_userinto="�b���򪺤j�D�W,%%�M�ۦۤv�a������j�¤l�A�⮳�@��j��M�Ө�F"&Application("ljjh_chatroomname")&"�A����B�کȽ֧r�I������F���򲳦�S��,�榣�����:���A�U��S���A�p�̦�§!��"
case 3
ljjh_userinto="�b���򪺤j�D�W,%%�M�ۤ@�Y�¤��j��ۤp���A�y���@��ͤF¸���}�K�C�A�Ө�F"&Application("ljjh_chatroomname")&",�����������r�I���A�����:���U��S�̡A�p�ͦ�§!��"
case 4
ljjh_userinto="�b���򪺤j�D�W,%%�M�ۤ@�ǿa�}�Ѱ��A�y���@��j�M�A�Ө�F"&Application("ljjh_chatroomname")&",��������ڦ�r�I���C�C���A�����:���U��S�̡A�p�ͦ�§!��"
case 5
ljjh_userinto="�b���򪺤j�D�W,%%�M�ۤ@�ǰ��Y�j���A�y���@����K���l�]�Ѧ���j�^�A�Ө�F"&Application("ljjh_chatroomname")&",�����C�C�`��{�N�ƤF�I���A�����:���U��S�̡A�p�ͦ�§!��"
case 6
ljjh_userinto="�b���򪺤j�D�W,%%�M�ۤ@���겣�}��O�]�·Ϫ��_�^�A�y���@�⤭4�A�Ө�F"&Application("ljjh_chatroomname")&",�����C�C�Ѥl�]�����ѡI���A�����:���U��S�̡A�p�ͦ�§!��"
case 7
ljjh_userinto="�b���򪺤j�D�W,%%�}�ۤ@���L�R�]�����I�I�I���ݮ�^�A�y���@��b�۰ʡA�Ө�F"&Application("ljjh_chatroomname")&",�����C�C�ڦ��֤F�I���A�����:���U��S�̡A�p�ͦ�§!��"
case 8
ljjh_userinto="�b���򪺤j�D�W,%%�}�ۤ@���L�R�]�����I�I�I���ݮ�^�A�y���@��b�۰ʡA�Ө�F"&Application("ljjh_chatroomname")&",�����C�C�ڦ��֤F�I���A�����:���U��S�̡A�p�ͦ�§!��"
case 9
ljjh_userinto="�b���򪺤j�D�W,%%�}�ۤ@�����}�]�u�O�ΪA�^�A�y���@����۰ʡA�Ө�F"&Application("ljjh_chatroomname")&",�����C�C�o�׬O�Ѥj�I���A�����:���U��S�̡A�p�ͦ�§!��"
case 10
ljjh_userinto="�b���򪺤j�D�W,%%�}�ۤ@���j�b�]���A�u�O�ΪA�A�𦺧A�I�^�A�y���Z���h�h�A�Ө�F"&Application("ljjh_chatroomname")&",�����C�C�o�׬O�u�����Ѥj�I���A�����:���U��S�̡A�p�ͦ�§!��"
end select
if info(4)>0 then
ljjh_userinto="�b���򪺤j�D�W,%%�}�ۤ@�����˪��_���⨮�]���A�����]�ˡA�����^�A�y�����ۦU���|���d�B���W�˺��F���ȯ]�_�A�Ө�F"&Application("ljjh_chatroomname")&"�A���A���|���N�O�n��,����F���򪺥S�̡̭A�޵۲����s�졧�O�S�ڡA�p�ߧڵ��A��|���d��"
end if
if info(6)="�x��" then
ljjh_userinto="�b���򪺤j�D�W,%%�}�ۤ@���겣�}��O�]�·Ϫ��_�^�A�y���@�⤭4�A�����ۨ�ӳC�Τj�~�A�Ө�F"&Application("ljjh_chatroomname")&"�A����B���O�x�����Ƚ֧r���C����F��������̤l,��:���p���̡A�����������ưȻ��ֻ���!���٫�ۥh�}�ɬ��|�O��"
end if
if info(0)=Application("ljjh_admin") then
ljjh_userinto="�b�F�C�����`���������W�A����%%�����y�F�C�Ծ��z�q�Ť��}�}�����A�@�r���O20�l�ӭ��椤������ʬ[�����C��Ѩ��U�F�C�Ծ��V�B�̴ͭ�����A�M��ۨ�G���ڨӤF,�j�a���ƽЧֻ�!��"
end if
if info(5)="�q�r��" then
	ljjh_userinto="ĵ���A����q�r��%%�w�g��J�ڭ̤����A���Q�@�@�Ǥ��i�i�H���ķ�A�U���j�L�p�߬��n   ��"
end if

if info(13)<>"�L" then
sd(199)="<font color=black>�i���i�j</font><font color=008800>" & Replace(ljjh_userinto,"%%","<img src='../ico/"& info(10) &"-2.gif' width='84' height='84'><a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('["& info(9) &"]') target=f2>" & info(0) &"</a>["&info(5)&"]["&info(6)&"]<font color=red><b>id:"& info(9) &"</b></font>�a�ۦۤv���Ĥl["&info(13)&"] <img src='../ico/"& info(10) &"-2.gif' width='16' height='16'>") & "</font><font class=t>(" & t & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
else
sd(199)="<font color=black>�i���i�j</font><font color=008800>" & Replace(ljjh_userinto,"%%","<img src='../ico/"& info(10) &"-2.gif' width='84' height='84'><a href=javascript:parent.sw('[" & info(0) & "]');parent.sws('["& info(9) &"]') target=f2>" & info(0) &"</a>["&info(5)&"]["&info(6)&"]<font color=red><b>id:"& info(9) &"</b></font>") & "</font><font class=t>(" & t & ")</font><bgsound src='readonly/cd.mid' loop='1'>"
end if

sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
'��end if�P�W���������i�J���i����
end if
end if
if info(1)="�k" then
if InStr(LCase(Application("ljjh_useronlinename"&session("nowinroom")))," " & LCase(pe) & " ")<>0 then
mess="�ѱC�j�H�ӤF�A�٤��֥h�Цw�I"
call messages(mess,pe)
end if
end if
sub messages(mess,pe)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=1
sd(194)="����"
sd(195)=pe
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i�t�Ρj"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
sub messages2(mess)
Application.Lock
sd=Application("ljjh_sd")
line=int(Application("ljjh_line"))
Application("ljjh_line")=line+1
for i=1 to 190
sd(i)=sd(i+10)
next
sd(191)=line+1
sd(192)=-1
sd(193)=1
sd(194)="����"
sd(195)=info(0)
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i�t�δ��ܡj"&mess&"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
chatroomname=Application("ljjh_chatroomname"&session("nowinroom"))
for i=0 to chatroomnum
ljchat=split(ljjh_roominfo(i),"|")
roomlist=roomlist&"<img src='room.gif'><a href=\"&chr(34)&"goroom1.asp?roomsn="&i&"&chatroomname="&ljchat(0)&"\"&chr(34)&" target=\"&chr(34)&"d\"&chr(34)&" title=\"&chr(34)&"  "&ljchat(3)&"\"&chr(34)&">"&ljchat(0)&"</a>"
next
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html;charset=big5">
<style>BODY{background-image: URL(bj.gif);background-position:'top right';background-repeat: no-repeat;background-attachment: fixed;}\n");</style>
<script Language="Javascript">
//self.resizeBy(screen.height,screen.width);
if(window!=window.top)
{
window.alert("�Шϥ�ie�s�����ϥΥ��t�ΡI");
top.location.href="../exit.asp"
}
if(window.name!="xajh2004")
{ var i=1;
while (i<=50)
{
window.alert("�A�Q�@����r�A�§ڡH�o�̬O���檺�A�h�O�B���h�a�I���I�C�C�I50���I�I");
i=i+1;
}
top.location.href="READONLY/BOMB.HTM"
}
var crm='<%=chatroomname%>';
var myn='<%=info(0)%>';
var cs=<%=info(2)%>;
var lst=0;
var tbclu=true;
var mdcls=true;
var listfaces=false;
var showhao=false;
var showpy="<%=info(11)%>";
var showmp="<%=info(5)%>";
var showsex="<%=info(1)%>";
var showseek="";
var chatcls = 0 ;
var clsok=0 ;
var bc = 1 ;
var showtype = 0 ;
headhigh=20;
var badword = new Array("�g��","�l","�h��","�Y��","�A��","�A�Q","��A","�u","�ާA","��") ;
var badstr = "~!@ #$%^&*()[]{}_+-|=\`;,:'\"?<>/����������������������������������������?������������������" ;

document.write("<title>"+crm+"</title>");
function write(cls){var fsize,lheight;
if(cls==1){fsize=this.f2.document.af.fs.value;lheight=this.f2.document.af.lh.value;}else{fsize='9';lheight='120';}
this.f1.document.open();
this.f1.document.writeln("<html><head><title>��ܰ�</title><meta http-equiv=Content-Type content=\"text/html; charset=big5\">");
this.f1.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:125%}.t{color:FF00FF;font-size:11pt;}body{font-family:\"�s�ө���\";font-size:11pt}A{text-decoration:none;color:FFFF00}A:Hover{text-decoration:underline;color:FFFF00}A:visited{color:FFFF00}BODY{background-image: URL('1c.gif');background-position:'right top';background-repeat: no-repeat;background-attachment: fixed;overflow-x:hidden;}</style></head>");
this.f1.document.writeln("<\script language=javascript>");
this.f1.document.writeln("var autoScroll=1;var scTimer;speedscroll=0");
this.f1.document.writeln("function scrollit(){if(!parent.f2.document.af.as.checked){autoScroll=0;}else{autoScroll=1;autoscrollnow();}return true;}");
this.f1.document.writeln("function speed(){if(!parent.f2.document.af.sp.checked){speedscroll=0;}else{speedscroll=1;}}");
this.f1.document.writeln("function autoscrollnow(){if (speedscroll==0 && autoScroll==1){this.scroll(0,65000);parent.f0.scroll(0,65000);setTimeout('autoscrollnow()',200);}else{var f1t=document.body.scrollTop;var f0t=parent.f0.document.body.scrollTop;var f1h=document.body.scrollHeight;var f0h=parent.f0.document.body.scrollHeight;if(autoScroll==1){if(f1t<=f1h||f0t<f0h){f1t+=15;f0t+=15;this.scroll(0,f1t);parent.f0.window.scroll(0,f0t);scTimer=setTimeout(\'autoscrollnow()\',10);}}}}<\/script>");
this.f1.document.writeln("<body oncontextmenu=self.event.returnValue=false background='../1.jpg' bgproperties='fixed' bgcolor=000000 text=DFF0C0>" );
this.f1.document.writeln("<span class=l><font color=red>�i��s�e���j</font>���P�w��<font color=red> "+myn+" </font>�Ө�m"+crm+"�n</span><br><\script language=javascript>autoscrollnow();<\/script>");
this.f1.document.writeln("<%=roomlist%><br><\script language=javascript>autoscrollnow();<\/script>");
this.f0.document.open();this.f0.document.writeln("<html><head><title>�������</title><meta http-equiv=Content-Type content=\"text/html; charset=big5\">");this.f0.document.writeln("<style type=text/css>.p{font-size:20pt}.l{line-height:" + lheight + "%}.t{color:FF00FF;font-size:11pt;}body{font-family:\"�s�ө���\";font-size:11pt;}A{text-decoration:none;color:FFFF00}A:Hover{text-decoration:underline;color:FFFF00}A:visited{color:FFFF00}body{font-family:\"�s�ө���\";font-size:11pt;}A{text-decoration:none}A:Hover{text-decoration:underline}A:visited{color:blue}BODY{background-image: URL('c2.gif');background-position:'right top';background-repeat: no-repeat;background-attachment: fixed;overflow-x:hidden;}</style></head>");this.f0.document.writeln("<body oncontextmenu=self.event.returnValue=false background='../2.jpg' bgproperties='fixed' bgcolor=000000 text=DFF0C0>");this.f0.document.writeln("<span class=l><font color=red>�i��s�e���j</font>���P�w��<font color=red> "+myn+" </font>�Ө�m"+crm+"�n�[�J�|�����I�k�U�����|���I</span><br>");this.t.location.href="t.asp";
}
function sh(n1id,n2id,mm,n1,n2,y1,y2,bq,ll){
chatcls=chatcls+1;
if (chatcls>400 && clsok==0)
{if(confirm("�A���̹���ܵo���`��Ƥw�g�W�L400��,�Ҽ{��z���q���ʯ���D,�t�Φ۰ʲM�̡A�[�ֲ�ѳt��,�O�_�M�̡H�I�����T�w���M�̡A�I�����������H�ᤣ�A�X�{������")){chatcls=0;parent.qp();}else{clsok=1;}}
var faces;
if (this.f2.document.af.soundtx.checked != true){ll=ll.replace(".wav","");ll=ll.replace(".mid","")}
if (listfaces==true){
faces="<img src='../ico/"+ this.f2.document.af.mdtx.value +"-2.gif' width='20' height='20'>"
}
else
{
faces="";
}
var show;
if (n2=="mingdan" & this.f2.document.af.listname.checked==true){n2="�j�a";if(parent.m.location.href=="about:blank"){parent.m.location.href="f3.asp";}else{parent.m.location.reload();}}
if(n1id==-1){
if(mm==1){show="<span class=l>�i�p��j"+ll+"</span><br>";}
else{show="<span class=l>"+ll+"</span><br>";}
}
else{
if(mm==1){show="<span class=l>�i�p��j";}
else{show="<span class=l>";}
if(n1==myn){show=show+faces+"<font color="+y1+">�A</font>"+bq;}
else{show=show+"<a href=javascript:parent.sw('["+n1+"]');parent.sws('["+n1id+"]') target=f2><font color="+y1+">"+n1+"</font></a>"+bq;}
if(n2==myn){show=show+"<font color=red> <font color="+y1+">"+n2+"</font> </font>";}
else{show=show+"[<font color="+y1+">"+n2+"</font>]";}
show=show+"���G"+"<font color="+y2+">"+ll+"</font></span><br>";}
if ((n1==myn | n2==myn)&tbclu==false){show="<div  style='background:#AAE9A3;'>"+show+"</div>"}
if(tbclu==true & (myn==n1 | myn==n2)){this.f1.document.writeln(show);}else{this.f0.document.writeln(show);}
}
function qp(){this.f0.location.href="about:blank";
this.f1.location.href="about:blank";setTimeout('parent.write(1)',500);}
function sw(username){var usna;usna=username.substring(1,username.length-1);this.f2.document.af.towho.options[0].value=usna;this.f2.document.af.towho.options[0].text=usna;this.f2.document.af.sytemp.focus();return;}
function sws(username){var usna;usna=username.substring(1,username.length-1);this.f2.document.af.towhoid.value=usna;this.f2.document.af.sytemp.focus();return;}

function md1(){this.f3.document.writeln("<html><head><meta http-equiv='content-type' content='text/html; charset=big5'><title>�b�u�Τ�C��</title><style type='text/css'>");
this.f3.document.writeln("body{font-family:�s�ө���;font-size:11pt;}td{font-family:�s�ө���;font-size:11pt;line-height:125%;}A{color:white;text-decoration:none;}A:Hover{color:yellow;text-decoration:none;}A:Active {color:ffcc00}</style>");
this.f3.document.writeln("<\script language=\"JavaScript\"><\/script>");
this.f3.document.writeln("</head><body oncontextmenu=self.event.returnValue=false  bgcolor=#A4530B  background=bk.jpg bgproperties=\"fixed\">");
}

function md2(){
var nowroom='<%=session("nowinroom")%>';
if (nowroom == 0){
this.f3.document.writeln("<br><a href=javascript:parent.sw('[�j�a]');parent.sws('[0]')>^_^�j�a^_^</a><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[�W�ťX��]]');parent.sws('[0]')><s>�W�ťX��]</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[�ԯ�]');parent.sws('[0]')><s>�ԯ�</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[�g���j�s]');parent.sws('[0]')><s>�g���j�s</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
this.f3.document.writeln("<a href=javascript:parent.sw('[�A�E���M]');parent.sws('[0]')><s>�A�E���M</s></a>&nbsp<font size=2 color=ffff00>NPC</FONT><br>");
}
if (nowroom == 1){
this.f3.document.writeln("<br><a href=javascript:parent.sw('[�j�a]');parent.sws('[0]')>^_^�j�a^_^</a><br>");
for(var gw=180;gw<390;gw++){gw=gw+30;
this.f3.document.writeln("<a href=javascript:parent.sw('["+gw+"�ũǪ�]');parent.sws('[0]') title=\"============&#13&#10�Ǫ����r ���Ŷå�&#13&#10============\";><font size=3 color=#FFFF00><img src='../ico/"+gw+".gif'  border='0' width='36' height='36'><s>"+gw+"�ũǪ�</s></font></a>&nbsp;<br>");}
for(var gw=490;gw<780;gw++){gw=gw+50;
this.f3.document.writeln("<a href=javascript:parent.sw('["+gw+"�ũǪ�]');parent.sws('[0]') title=\"============&#13&#10�Ǫ����r ���Ŷå�&#13&#10============\";><font size=3 color=#FFFF00><img src='../ico/"+gw+".gif'  border='0' width='36' height='36'><s>"+gw+"�ũǪ�</s></font></a>&nbsp;<br>");}
for(var gw=801;gw<1100;gw++){gw=gw+49;
this.f3.document.writeln("<a href=javascript:parent.sw('["+gw+"�ũǪ�]');parent.sws('[0]') title=\"============&#13&#10�Ǫ����r ���Ŷå�&#13&#10============\";><font size=3 color=#FFFF00><img src='../ico/"+gw+".gif'  border='0' width='36' height='36'><s>"+gw+"�ũǪ�</s></font></a>&nbsp;<br>");}
this.f3.document.writeln("<a href=# onClick=window.open('hqg.asp','hqg','width=500,height=400');><br><img src='pic/h0.gif'  border='0'><font color=yellow><font size=3>�x�C��</font></font></a>&nbsp;<br>");}
}
function tma(onlineok){
var ss;
var faces;
var myself;
var friend;
var colorok;
var hycolor;
ljjh_zt=onlineok.split("|")
if (showpy.search(ljjh_zt[2]) != -1){friend="<img src='../jhimg/friend.gif' width='12' height='12'>";}
else
{friend=""}
if (ljjh_zt[2]==myn){myself="<img src='../jhimg/self.gif' width='12' height='12'>";}
else
{myself="";
}
if (listfaces==true){
faces="<img src='../ico/"+ljjh_zt[8]+"-2.gif' width='"+headhigh+"' height='"+headhigh+"'>"+friend+myself;
}
else
{
faces="";
}
if(ljjh_zt[3]=="�k"&&ljjh_zt[0]==0){hycolor="#DEECFD";faces="<img src='boy.gif'>";}
if(ljjh_zt[3]=="�k"&&ljjh_zt[0]>0){hycolor="#20AFE0";faces="<img src='boy.gif'>";}
if(ljjh_zt[3]=="�k"&&ljjh_zt[0]==0){hycolor="#FFD0CF";faces="<img src='girl.gif'>";}
if(ljjh_zt[3]=="�k"&&ljjh_zt[0]>0){hycolor="#F04F00";faces="<img src='girl.gif'>";}
if(ljjh_zt[1]=="�x��"){colorok="#7FCF50";faces="<img src='room.gif'>";}else colorok="#FFFF00";
if(ljjh_zt[5]=="�x��"){colorok="#FFC01F";faces="<img src='room.gif'>";}
if(ljjh_zt[1]=="����"){colorok="8FD060";faces="<img src='room.gif'>";}
ss=faces+"<a href=\"javascript:parent.sw(\'[" + ljjh_zt[2] + "]\');\"";
ss=ss+"><font size=2 color="+hycolor+">"+ljjh_zt[2]+"</font></a>"+" <font size=2 color="+colorok+">"+ljjh_zt[1]+"</font><br>";
//������ܦn�ͪ��B�z
this.f3.document.writeln(ss);
}
function mdsm(){
if (listfaces==false){
this.f3.document.writeln("<br></table>");
;}
}
function md3(){this.f3.document.writeln("</td></tr></table></body></html>");this.f3.document.close();}
function fc(){rn();setTimeout('parent.write()',0);}
function rn(){this.f2.document.af.username.value=myn;}
function clsay(){if(cs<6){setTimeout("this.f2.document.af.sytemp.value=''",0);}}
function IsBadWord(m)
{var tmp = "" ;
for(var i=0;i<m.length; i++)
{for(var j=0;j<badstr.length;j++)
if(m.charAt(i) == badstr.charAt(j)) break;
if(j==badstr.length) tmp += m.charAt(i) ;}
for(i=0;i<badword.length;i++) if(tmp.search(badword[i]) != -1) return true;
return false;}
function Warning()
{this.f2.document.af.sytemp.value='';
if(bc > 3) top.location.href="autokick.asp";else{bc++;alert("�Ф��n�b��ѫǤ���ż��,�_�h��X�I�I�I");}}
function checksays()
{if(this.f2.document.af.addvalues.checked)
{alert("�z�ثe���}�F�۰ʪw�I�\��A����o���I")
return false;}
if(this.f2.document.af.sytemp.value=="/���u$"){window.open("maninfosearch.asp?id="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(this.f2.document.af.sytemp.value=="/�d�ݪ��~$"){window.open("towupin.asp?toname="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(this.f2.document.af.sytemp.value=="/�H��$"){window.open("../yamen/mt.asp?action="+this.f2.document.af.towho.value);this.f2.document.af.sytemp.value="";return false;}
if(IsBadWord(this.f2.document.af.sytemp.value)){Warning();return false;}
this.f2.document.af.sy.value='';if(this.f2.document.af.sytemp.value!='')
{
this.f2.document.af.oldtowho.value=this.f2.document.af.towho.value;
this.f2.document.af.sy.value=this.f2.document.af.sytemp.value;
this.f2.document.af.command.options[0].selected=true;
this.f2.document.af.command2.options[0].selected=true;
this.f2.document.af.command3.options[0].selected=true;
this.f2.document.af.sytemp.focus();
this.f2.document.af.sytemp.value='';ty=new Date();var nh=ty.getHours();var nm=ty.getMinutes();var ns=ty.getSeconds();var ct=(nh*3600)+(nm*60)+ns;if((ct-lst)>2.5){lst=ct;}else{this.f2.af.sytemp.value=this.f2.af.oldsays.value;
this.f2.af.oldsays.value='';return false;}this.f2.addOne(this.f2.document.af.sy.value);
this.f2.startnosay();
this.f2.document.af.subsay.disabled=1;setTimeout('this.f2.document.af.subsay.disabled=0',3000);
return true;}if ((this.f2.document.af.zt.options[this.f2.document.af.zt.selectedIndex].value=='0')||(this.f2.document.af.zt.options[this.f2.document.af.zt.selectedIndex].value=='')){alert('�п�J�o���ο�ܰʧ@�I');
this.f2.document.af.sytemp.focus();
this.f2.document.af.sytemp.select();return false;}}
function tbclutch(){
if(this.f2.document.af.tbclutch.value=='����')
{this.f2.document.af.tbclutch.value='����';
this.msgfrm.rows = "*";
this.msgfrm.cols = "*";
tbclu=false;}
else
{
if(this.f2.document.af.tbclutch.value=='����')
{this.f2.document.af.tbclutch.value='����';
this.msgfrm.cols = "*,*";
this.msgfrm.rows = "*";
tbclu=true;}
else
{
this.f2.document.af.tbclutch.value='����';
this.msgfrm.cols = "*";
this.msgfrm.rows = "*,*";
tbclu=true;}}
this.f2.document.af.sytemp.focus();}
function face(){
if(this.f2.document.af.listface.checked==true){
parent.m.location.reload();
listfaces=true;
}
else{
parent.m.location.reload();
listfaces=false;
}
this.f2.document.af.sytemp.focus();
}
self.onerror=null;
var nullframe = '<HTML><BODY BGCOLOR=#000000 text=#ffffff><center><H3 color=yellow><font color=yellow>�w��Ө��F�C����</font></h3><br><br>�{�ǥ��b�q�A�Ⱦ�Ū�����, �еy�� ......</center></BODY></HTML>';
</script>
</head>
<frameset name=mainfrm cols="*" rows="*,0" border=0>
<frameset cols="*,154" name=tbgn1 rows="*" border="0" framespacing="0" frameborder="NO">
<frameset rows="*,0,70,0" cols="*">
<frameset name=msgfrm rows="*,*" cols="*">
<frame src="javascript:parent.nullframe" name="f0" scrolling="AUTO" marginheight="3" marginwidth="5" style="border-top:1px #9A87DE dashed;border-left:1px #9A87DE dashed">
<frame src="javascript:parent.nullframe" name="f1" scrolling="AUTO" marginheight="3" marginwidth="5" style="border-top:1px #9A87DE dashed;border-left:1px #9A87DE dashed">
</frameset>
<frame src="about:blank" name="t" marginwidth="5" marginheight="5" noresize scrolling="NO">
<frame src="f2.asp" name="f2" scrolling="NO" marginwidth="3" marginheight="8" noresize  style="border-top:1px #9A87DE dashed;border-left:1px #9A87DE dashed">
<frame src="option.asp" name="d" scrolling="NO">
</frameset>
<frameset rows="0,510,0,150" cols="*" name=tbymd>
<frame src="f3.asp" name="m" noresize scrolling="NO">
<frame src="about:blank" marginwidth="5" marginheight="5" name="f3">
<frame src="about:blank" marginwidth="5" marginheight="5" name="ps">
<frame src="f4.asp" scrolling="NO" noresize name="f4" marginwidth="3" marginheight="3">
</frameset>
</frameset>
<frame name="optfrm" src="about:blank" marginheight=0 marginwidth=0 scrolling=no noresize>
</frameset><noframes></noframes></html>
