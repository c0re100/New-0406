<%@ LANGUAGE=VBScript codepage ="950" %>
<%Response.charset="big5"%>
<%Response.Buffer=true
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = 0
if not IsArray(Session("info")) then Response.Redirect "../error.asp?id=440"
info=Session("info")
id=LCase(trim(request.querystring("id")))
yn=LCase(trim(request.querystring("yn")))
if InStr(id,"or")<>0 or InStr(id,"'")<>0 or InStr(id,"`")<>0 or InStr(id,"=")<>0 or InStr(id,"-")<>0 or InStr(id,",")<>0 then 
	Response.Write "<script language=JavaScript>{alert('�u�a�A�A�Q������H�Q�o�öܡH�I');}</script>"
	Response.End 
end if
too=trim(Application("ljjh_hunyin"))
Set conn=Server.CreateObject("ADODB.CONNECTION")
Set rs=Server.CreateObject("ADODB.RecordSet")
conn.open Application("ljjh_usermdb")
rs.open "select �m�W FROM �Τ� WHERE id="&id,conn
peiou=rs("�m�W")
'if info(0)<>trim(Application("ljjh_hunyin")) then
'	rs.close
'	set rs=nothing
'	conn.close
'	set conn=nothing
'	Response.Write "<script language=JavaScript>{alert('�A�Q�@����A�H�a�]�S�������A�I');}</script>"
'	Response.End
'end if
if yn=1 then
if info(0)=too then
	Response.Write "<script language=JavaScript>{alert('�ڵ����A���n�D�F�I');}</script>"
	conn.execute "update �Τ� set �t��='"& peiou &"',���B�ɶ�=date() where id="&info(9)
	conn.execute "update �Τ� set �t��='"& info(0) &"',���B�ɶ�=date() where id="&id
	hunyin="���ߡG["& info(0) &"]�P{"& peiou &"}�ߵ��}�t�A�j�a��ܯ��P�I�I<img src='img/004.gif'><br><marquee width=100% behavior=alternate scrollamount=5><font color=red size=+1>�߳�</font></marquee>"
   else
  randomize()
		r = Int(7 * Rnd)+1
	Select Case r
		Case 1
			hunyin=info(0)&"���G����G"&info(0)&"�ڤ��["&peiou&"�P"&too&"]���B�A�@�w�檺   "
		case 2 
			hunyin=info(0)&"���G����G"&info(0)&"�ڤ��["&peiou&"�P"&too&"]���B,�A�O�Ƥl�L�O��,�G�b�ݤl�֥b��,�S���A�N���ƨ�,�֤F�L�]�N���b��,�A�b���ӥL�b��,�A�n���F�N�֭�."
		case 3
			hunyin=info(0)&"���G����G"&info(0)&"�ڤ��["&peiou&"�P"&too&"]���B,��M�A�����C�����E�ٴ��A�H�~�ϴ_��C�۫�ۨ������A���ɦ��]�������C�o�بk�H�w�g����F�A�A�N�����L�a!"
		case 4
			hunyin=info(0)&"���G����G"&info(0)&"�ڤ��["&peiou&"�P"&too&"]���B,�ڤӨتA�A�F�A�ӥB���o�򻡪��@�w�O�i�����H�C�ڥ��ɤ��S�ɡA�@���첧�ʴN�y���A�u�S��k�C"
		case 5
			hunyin=info(0)&"���G����G"&info(0)&"�ڤ��["&peiou&"�P"&too&"]���B,����@�����ݳz�A�Ʊ�o�೭�A�ݲӤ��`�y�I�p�G�o�������A,���i��O�o�Q�M�ڬݡC"
		case else
			hunyin=info(0)&"���G����G"&info(0)&"�ڤ��["&peiou&"�P"&too&"]���B�A�ѯ�a�ѡA�Ʊ�A�̱����� "
	end select
	'if peiou=info(0) then
	'		hunyin=info(0)&"���G����G"&info(0)&"�ڤ��["&peiou&"�P"&too&"]���B�A�u�O���n�y�A���̦��ۤv���ۤv�n���I"
	'end if
	end if

else
if info(0)=too then
	Response.Write "<script language=JavaScript>{alert('�ڤ~���z�O�I');}</script>"
	conn.execute "update �Τ� set �y�O=�y�O-100 where id="&id
	randomize()
		r = Int(6 * Rnd)+1
		Select Case r
			Case 1
				hunyin="�i�ڱB�j"&info(0)&"��"&peiou&"���G�A�i�H����ڧ�n���H  �ڤ��Q�A���F�A�M�ۤv�A�Ʊ�A��̧ڧa�I�ڷ|�û����֧A��!["& peiou &"]�V{"& info(0) &"}�D�B�����A�y�O�U�N100�I!"
			case 2
				hunyin="�i�ڱB�j"&info(0)&"��"&peiou&"���G�ڮa�̪������ӳ��w�A�A�A�a�̪��ߤ��ӳ��w�ڡA�ҥH���n�A������A�H��A�C�C�N���ժ�!"
			case 3
				hunyin="�i�ڱB�j"&info(0)&"��"&peiou&"���G������ɡA�R�������Ʀ�A���h�w�L�k���^.....���B�ͦn�ܡH"
			case 4
				hunyin="�i�ڱB�j"&info(0)&"��"&peiou&"���G���g���ɮ��w�g�h��A�ڪ��߸̤w�g���A���E���A�ЧA�٬O��O���H�a�C"
			case else
				hunyin="�i�ڱB�j"&info(0)&"��"&peiou&"���G�A�ݬݧڪ��ˡA�A�]�t�A�ڴN�O�����޵L��A�]���|�����A�o��ž�T "
			end select
'hunyin="�i�ڱB�j"&info(0)&"��"&peiou&"���G�A�i�H����ڧ�n���H  �ڤ��Q�A���F�A�M�ۤv�A�Ʊ�A��̧ڧa�I�ڷ|�û����֧A��!["& peiou &"]�V{"& info(0) &"}�D�B�����A�y�O�U�N100�I!"
else
		randomize()
		r = Int(6 * Rnd)+1
	Select Case r
		Case 1
			hunyin=info(0)&"���G�Ϲ�G"&info(0)&"�ڤϹ�["&peiou&"�P"&too&"]���B�A�L�̮ڥ����X�A   �٤��p�ڡI"
		case 2
			hunyin=info(0)&"���G�Ϲ�G"&info(0)&"�ڤϹ�["&peiou&"�P"&too&"]���B�A�O�����L�A�A�ݧڤ�L�^�T�h�F�A�ڦ����l�A�Фl�A���l�A�{�b�N�ʩd�l�F�C�M��ZZZ�J�J�L�Q�G����~~~���F�F�A�A��~~~"
		case 3
			hunyin=info(0)&"���G�Ϲ�G"&info(0)&"�ڤϹ�["&peiou&"�P"&too&"]���B�A�L�O��T�I�e�X�ѧ��٬ݨ�L�q�R�K�|�X��  ����I�I"
		case 4
			hunyin=info(0)&"���G�Ϲ�G"&info(0)&"�ڤϹ�["&peiou&"�P"&too&"]���B�A�L�O��T�I�e�X�ѧ��٬ݨ�L�q�R�K�|�X��  ����I�I"
		case 5
			hunyin=info(0)&"���G�Ϲ�G"&info(0)&"�ڤϹ�["&peiou&"�P"&too&"]���B�A�L�걡���N ���Ȱ��|�簨�� !"
		case else
			hunyin=info(0)&"���G�Ϲ�G"&info(0)&"�ڤϹ�["&peiou&"�P"&too&"]���B�A�����L�e�{�����R���ơA�����L��K�����ܬ����c�F�s�����}���u���A�]�ܦ��F�걡���N�����F�A�ҥH�ЧA�d�U���n�����L!"
	end select
	end if
end if
rs.close
set rs=nothing
conn.close
set conn=nothing
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
sd(196)="B0D0E0"
sd(197)="B0D0E0"
sd(198)="��"
sd(199)="<font color=B0D0E0>�i��������j</font>"&hunyin
sd(200)=session("nowinroom")
Application("ljjh_sd")=sd
'Application("ljjh_hunyin")=""
Application.UnLock
%>
 