<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=false
if not IsArray(Session("info")) then Response.Redirect "../../error.asp?id=440"
info=Session("info")
if Session("ljjh_inthechat")<>"1" then %>
<script language="vbscript">
alert "�A����i��ާ@�A�i�榹�ާ@�����i�J��ѫǡI"
location.href = "javascript:history.go(-1)"
</script>
<%end if
ljjh_roominfo=split(Application("ljjh_room"),";")
chatinfo=split(ljjh_roominfo(session("nowinroom")),"|")
num=trim(request.form("num"))
for t=1 to len(num)
abc=mid(num,t,1)
if abc<>"0" and abc<>"1" and abc<>"2" and abc<>"3" and abc<>"4" and abc<>"5" and abc<>"6" and abc<>"7" and abc<>"8" and abc<>"9" then
	Response.Write "<script language=JavaScript>{alert('�ާ@���~�A�U�`�ШϥμƦr�I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end if
next
yinliang=abs(int(num))
if Request.Form("submit")="�j" then
tt="�j"
end if
if Request.Form("submit")="�p" then
tt="�p"
end if
if Request.Form("submit")="��" then
tt="��"
end if
if Request.Form("submit")="��" then
tt="��"
end if
select case tt
	case "�j"
		duboimg="<img src='../jhimg/da.gif'>"
	case "�p"
		duboimg="<img src='../jhimg/xiao.gif'>"
	case "��"
		duboimg="<img src='../jhimg/dan.gif'>"
	case "��"
		duboimg="<img src='../jhimg/shuang.gif'>"
case else
	Response.Write "<script language=JavaScript>{alert('��ާ@���G�j�B�p�B��B���I');location.href = 'javascript:history.go(-1)';}</script>"
	Response.End 
end select
if yinliang<10 or yinliang>500 then
		Response.Write "<script language=JavaScript>{alert('�̤֩�G10�I�԰��g��A�̦h500�I�԰��g��I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
'�}�l�P�_
	Set conn=Server.CreateObject("ADODB.CONNECTION")
	Set rs=Server.CreateObject("ADODB.RecordSet")
	conn.open Application("ljjh_usermdb")
	rs.open "select allvalue FROM �Τ� WHERE id="&info(9),conn
	if rs("allvalue")<yinliang then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�A���o��h���s�I�ܡH�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.close
	tmprs=conn.execute("Select count(*) As �ƶq from �I�ƽ䧽 where �I��<>0")
	durs=tmprs("�ƶq")
set tmprs=nothing
	if durs>=5 then
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�@["&durs&"]�`���b���ݶ}���A�y��A�U�`�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.open "select top 1 �m�W FROM �I�ƽ䧽 WHERE ����='���a'",conn
	if rs.eof or rs.bof then
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('�{�b�S�����a�A����i���I�ƽ䧽�I�I');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 		
	end if
	rs.close
	rs.open "select top 1 * FROM �I�ƽ䧽 WHERE �m�W='"& info(0) &"'",conn
	if not(rs.eof) or not(rs.bof) then
		if rs("����")="���a" then
			temp=info(0)&"�A�{�b�O���a�r�A�A�n�d����r!"
		else
			temp=info(0)&"�A��["&rs("��j�p")&"]"&rs("�I��")&"���۶}�a�I"
		end if
		rs.close
		set rs=nothing	
		conn.close
		set conn=nothing
		Response.Write "<script language=JavaScript>{alert('"&temp&"');location.href = 'javascript:history.go(-1)';}</script>"
		Response.End 
	end if
	rs.close
	conn.execute "update �Τ� set allvalue=allvalue-"&yinliang&" where id="&info(9)
	conn.execute "insert into �I�ƽ䧽(�m�W,����,�I��,��j�p,�ɶ�) values ('"&info(0) &"','���a',"&yinliang&",'"&tt&"',now())"	
	tmprs=conn.execute("Select count(*) As �ƶq from �I�ƽ䧽 where ��j�p='�j'")
	dars=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �I�ƽ䧽 where ��j�p='�p'")
	xiaors=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �I�ƽ䧽 where ��j�p='��'")
	danrs=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �I�ƽ䧽 where ��j�p='��'")
	shuangrs=tmprs("�ƶq")
	tmprs=conn.execute("Select count(*) As �ƶq from �I�ƽ䧽 where �I��<>0")
	durs=tmprs("�ƶq")


'�}���F
if durs>=5 then
	Randomize
	m1 = Int(6 * Rnd + 1)
	Randomize
	m2 = Int(6 * Rnd + 1)
	Randomize
	m3 = Int(6 * Rnd + 1)
	sjdubo=m1+m2+m3
'�d����a
rs.open "select �m�W FROM �I�ƽ䧽 WHERE ����='���a'",conn
zhuangjia=rs("�m�W")
rs.close

'�\�l�B�z
if m1=m2 and m2=m3 and m3=m1 then
	rs.open "select �I�� FROM �I�ƽ䧽 WHERE �I��<>0 and ����<>'���a'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("�I��")
	rs.movenext
	loop
	rs.close
	conn.execute "update �Τ� set allvalue=allvalue+"&yingyin&" where �m�W='"& zhuangjia &"'"
	xiazhu="���a�}�G<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�p�G<font color=red>"& sjdubo &"</font>�I!���a�}�X�\�l  �q���I���a�G<font color=red>"&zhuangjia&"</font>Ĺ�o�G<font color=red>"&yingyin&"</font>�I�s�I�I"
		Application.Lock
	Application("ljjh_dubozhang3")=""
Application.UnLock
set rs=nothing
		conn.close
		set conn=nothing
		call say(xiazhu)
Response.Write "<script Language=Javascript>alert('�U�`"&yinliang&"�I�w�I���\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End
end if

'��l�Ƽƾ�
shuangyinliang=0
shuangname=""
danyinliang=0
danname=""
xiaoyinliang=0
xiaoname=""
dayinliang=0
daname=""

'�}�����B�z
if sjdubo/2=int(sjdubo/2) then
	danshuang="<img src='../jhimg/shuang.gif'>"
	rs.open "select �I��,�m�W FROM �I�ƽ䧽 WHERE ��j�p='��'",conn
	do while not rs.bof and not rs.eof
		shuangyinliang=shuangyinliang+rs("�I��")
		shuangname=shuangname&rs("�m�W")&" "
		conn.execute "update �Τ� set allvalue=allvalue+"&rs("�I��")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set allvalue=allvalue-"& rs("�I��") &" where �m�W='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from �I�ƽ䧽 where  ��j�p='��'"
else
	danshuang="<img src='../jhimg/dan.gif'>"
	rs.open "select �I��,�m�W FROM �I�ƽ䧽 WHERE ��j�p='��'",conn
	do while not rs.bof and not rs.eof
		danyinliang=danyinliang+rs("�I��")
		danname=danname&rs("�m�W")&" "
		conn.execute "update �Τ� set allvalue=allvalue+"&rs("�I��")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set allvalue=allvalue-"& rs("�I��") &" where �m�W='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from �I�ƽ䧽 where  ��j�p='��'"
end if


'��}�j�p�B�z
if sjdubo<=10 then
	daxiao="<img src='../jhimg/xiao.gif'>"
	rs.open "select �I��,�m�W FROM �I�ƽ䧽 WHERE ��j�p='�p'",conn
	do while not rs.bof and not rs.eof
		xiaoyinliang=xiaoyinliang+rs("�I��")
		xiaoname=xiaoname&rs("�m�W")&" "
		conn.execute "update �Τ� set allvalue=allvalue+"&rs("�I��")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set allvalue=allvalue-"& rs("�I��") &" where �m�W='"& zhuangjia &"'"
	rs.movenext	
	loop
	rs.close
	conn.execute "delete * from �I�ƽ䧽 where  ��j�p='�p'"

else
	daxiao="<img src='../jhimg/da.gif'>"
	rs.open "select �I��,�m�W FROM �I�ƽ䧽 WHERE ��j�p='�j'",conn
	do while not rs.bof and not rs.eof
		dayinliang=dayinliang+rs("�I��")
		daname=daname&rs("�m�W")&" "
		conn.execute "update �Τ� set allvalue=allvalue+"&rs("�I��")*2&" where �m�W='"& rs("�m�W") &"'"
		conn.execute "update �Τ� set allvalue=allvalue-"& rs("�I��") &" where �m�W='"& zhuangjia &"'"	
	rs.movenext
	loop
	rs.close
	conn.execute "delete * from �I�ƽ䧽 where  ��j�p='�j'"
end if

'��ѤU�骺�Τ��I��
	rs.open "select �I�� FROM �I�ƽ䧽 WHERE �I��<>0 and ����<>'���a'",conn
	yingyin=0
	do while not rs.bof and not rs.eof
	yingyin=yingyin+rs("�I��")
	rs.movenext
	loop
	rs.close
	conn.execute "update �Τ� set allvalue=allvalue+"&yingyin&" where �m�W='"& zhuangjia &"'"
	conn.execute "delete * from �I�ƽ䧽 where  �m�W<>''"
	zong=yingyin+shuangyinliang+danyinliang+xiaoyinliang+dayinliang
	pei=shuangyinliang+danyinliang+xiaoyinliang+dayinliang
'shuangyinliang=0
'shuangname=""
'danyinliang=0
'danname=""
'xiaoyinliang=0
'xiaoname=""
'dayinliang=0
'daname=""
	duboname=shuangname&danname&xiaoname&daname

	
	xiazhu="���a�}�G<img src='../jhimg/"&m1&".gif'><img src='../jhimg/"&m2&".gif'><img src='../jhimg/"&m3&".gif'>�p�G<font color=red>"& sjdubo &"</font>�I!"&danshuang&daxiao&"�`�U�`�G<font color=red>"&zong&"</font>�I�w�I�A���a�G<font color=red>["&zhuangjia&"]</font>���J�G<font color=red>"&yingyin &"</font>�I�w�I,�ߥX�G<font color=red>"&pei&"</font>�I�w�I�A�X�p�G<font color=red>"&yingyin-pei&"</font>�I�w�I,�@���G<font color=red>"&duboname&"</font>���a���B�㤤�I"
Application.Lock
	Application("ljjh_dubozhang3")=""
Application.UnLock
	set rs=nothing
	conn.close
	set conn=nothing
	call say(xiazhu)
Response.Write "<script Language=Javascript>alert('�U�`"&yinliang&"�I�w�I���\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End

end if
xiazhu="<a href=javascript:parent.sw('[" & info(0) & "]'); target=f2><font color=red>"&info(0)&"</font></a>�߯k�a��ۤv�V�O�w�Ӫ��s�I���X:<font color=red>"& yinliang &"</font>�I�A�ک�"& duboimg &"�I�A�@�w�����I�{�b�U�`�p�U�G��j�G<font color=red>"& dars &"</font>�ӡA��p:<font color=red>"& xiaors &"</font>��,���G<font color=red>"&danrs&"</font>��,����:<font color=red>"&shuangrs&"</font>�ӡI�ٮt:<font color=red>"&(5-durs)&"</font>�Ӷ}���I"

	set rs=nothing	
	conn.close
	set conn=nothing

	call say(xiazhu)
Response.Write "<script Language=Javascript>alert('�U�`"&yinliang&"�I�w�I���\�I');location.href = 'javascript:history.go(-1)';</script>"
Response.End

sub say(xiazhu)
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
sd(194)=info(0)
sd(195)=info(0)
sd(199)="<font color=green>�i�g��䧽�j</font><font color=#FFC01F>"& xiazhu &"</font>"
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
Application.UnLock
end sub
%>