$|=1;
#---------------------------------------------------#
$datadir    = "./book";
$user=$FORM{'user'};
$user = "bearcc" if ($user eq "");
$setup_path = "$datadir/$user\-setup.cgi";
$gbdata     = "$datadir/$user\-book.txt";
$data	    = "$datadir/$user\-data.txt";
$count      = "$datadir/$user\-count.txt";
$mail = "/usr/sbin/sendmail";
$picdir     = "./bookpic";	#��m���ɥؿ�
$kickurl    = "./bookpic";	#��m�ϧ����ɥؿ�
require "$setup_path";
#---------------------------------------------------#
$useimg = "OFF";
#---------------------------------------------------#
$job=$FORM{'job'};
$page=$FORM{'page'};
$searchword=$FORM{'searchword'};
$userip=$ENV{'REMOTE_ADDR'};
$del=0; $safe=0; $cot=0;
#---------------------------------------------------#
if ($job eq "add") {	&add; }
elsif ($job eq "post") { &post; }
elsif ($job eq "super") { &super; }
elsif ($job eq "admin") { &admin; $del=1; &look; }
elsif ($job eq "config") { &admin; $cot=1; &config; }
elsif ($job eq "safe") { &admin; $del=1; $safe=1; &look; }
elsif ($job eq "configadd") { &cfgadd; }
elsif ($job eq "del") { &del; }
else { &look; }
#---------------------------------------------------#
sub look {
open (FILE, $gbdata) || &error("�L�k�}��$gbdata��");@GBDATA=<FILE>;close (FILE);
open(COUNT,"$count") || &error("�L�k�}��$count��");@counters=<COUNT>;close(COUNT);
foreach (@counters) {
($lastip,$counter)=split("��");
if ($userip ne $lastip) {
$counter++;
open(COUNT,">$count") || &error("�L�k�}��$count��");flock(COUNT,2);print COUNT "$userip��$counter\n";flock(COUNT,8);close(COUNT);}}
$i=0;$j=1;$temp=1;$flag=0;
if ($searchword ne "") {
$searchword1=$searchword;$searchword=quotemeta $searchword;
foreach (@GBDATA) {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/��/,$_);
if ($chkmsg =~ /$searchword/i) { push(@temp,"$_"); }
elsif ($chkname =~ /$searchword/i) { push(@temp,"$_"); }
elsif ($chkmsg2 =~ /$searchword/i) { push(@temp,"$_"); }
}
@GBDATA=@temp;}
foreach $line (@GBDATA) {
if ($line =~ /<!--(.*)-->/) { $i++; } }
$den=($i%$show);
if ($den != 0 || $i == 0){ $den=int($i/$show)+1; }
else { $den=($i/$show); }
if ($page eq ""){ $min=1; $max=$show; $page=1; }
if ($page ne ""){ $min=(($page-1)*$show)+1; $max=$page*$show; }
$l=$i;
&top;&menu;
print <<AAA;
<td width="99%" valign="top" bgcolor="#ffffff" style="border-collapse: collapse ;border: 1px dashed  #000000">
<div align="center">
AAA
&start;
print <<AAA;
<table border=0 cellspacing=0 cellpadding=0 width=99%>
AAA
print "<tr><td align=left><select name=\"dennis\" onchange=\"location.href=this.options[this.selectedindex].value;\" style=\"background:$bgcolor;color:$text\">\n";
while ($l>0) {
if ($page ne $j) {
print "<OPTION VALUE=\"./?";
if ($del eq 1) { print "job=admin&"; }
print "action=aa7&user=$user&page=$j&searchword=$searchword1";
if ($del eq 1) { print "&id=$FORM{'id'}&passid=$FORM{'passid'}"; }
print "\">";
print "��";
if ($j <10 ){ print "0"; }
print "$j��";
print "</OPTION>\n";}
else { print "<option value='' selected>��";
if ($j <10 ){ print "0"; }
print "$j��*</OPTION>\n"; }
$l=$l-$show; $j++;}
$uptime  = `uptime`;
@uptime =split(/\s+/,$uptime);
$uptime = @uptime[$#uptime];
print "</select></td>\n";
if ($searchword ne "") {
print "<td align=center>��� <font color=red><b>$i</b></font> ��</td>";}
else {print "<td align=center>�@�� <font color=red><b>$i</b></font> ��</td>";}
print "<td align=center>�y�k�G<font color=red><b>$usehtml</b></font></td>";
print "<td align=center>�K�ϡG<font color=red><b>$useimg</b></font></td>";
print "<td align=center>�t�έt���G<font color=red><b>$uptime</b></font></td>";
print "<td align=center><img src=./count/a.gif>";

while(length($counter) < 6){ $counter = 0 . $counter; }
@cnts = split(//,$counter);
print "<img src=\"./count/$cnts[0]\.gif\">";print "<img src=\"./count/$cnts[1]\.gif\">";
print "<img src=\"./count/$cnts[2]\.gif\">";print "<img src=\"./count/$cnts[3]\.gif\">";
print "<img src=\"./count/$cnts[4]\.gif\">";
print "<img src=./count/b.gif></td></tr></table>\n";
foreach $line (@GBDATA) {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/��/,$line);
chop($line);
$chkcnt=~ s/<!--//g;
$chkcnt=~ s/-->//g;
$chkmsg =~s/.com/www/g;$chkmsg =~s/chat/www/g;
$chkmsg =~s/love/www/g;$chkmsg =~s/http/www/g;
$chkmsg =~s/hk/www/g;
if ($safe eq "1") { $chkmsg=~ s/</&lt;/g; $chkmsg=~ s/>/&gt;/g;	$chkmsg=~s/&lt;BR&gt;/<BR>\n/g; $chkmsg="<FONT COLOR=RED>���w���Ҧ���</FONT><BR>$chkmsg"; }
if ($del ne 1 && $chkhidd eq "1") { $chkmsg="<FONT COLOR=$text>�i�o�O��<B><FONT COLOR=RED>$manger</B></FONT>�������ܡD�D�D�D�j"; }
if ($del eq 1 && $chkhidd eq "1") { $chkmsg="<FONT COLOR=RED>�������ܤ��e��</FONT><BR>$chkmsg"; }
if ($del ne 1 && $chkhidd eq "2") { $chkmsg="$actext"; }
if ($del eq 1 && $chkhidd eq "2") { $chkmsg="<FONT COLOR=RED>���Q��������Ť��e��</FONT><BR>$chkmsg"; }
$chkmsg =~s/$searchword/<FONT COLOR=RED>$searchword1<\/FONT>/g if ($searchword ne "");
$chkmsg2=~s/$searchword/<FONT COLOR=RED>$searchword1<\/FONT>/g if ($searchword ne "") && ($chkmsg2);
$chkmsg=~s/<BR>/<BR>\n/g;
if ($temp>=$min && $temp <=$max) {
print "<table border=\"0\" cellspacing=\"0\" cellpadding=\"3\" width=\"99%\">\n";
print "<tr bgcolor=$tdcolor><td width=140>\n";
if ($chkdate eq $Date) {print "<img src=bookpic/new.gif>";}
print "<font color=#ffffff>�E$chkname�E</font></td>\n";
print "<td width=110>\n";
if (($chkwww) && ($chkurl) || $chkemail) {
if (($chkwww) && ($chkurl)){
print "<font color=ffffff><a href=\"$chkurl\" target=\"_blank\" title=\"$chkwww\"><font color=#ffffff>����</a><font color=ffffff></font>";
}
if ($chkemail) {
print "<font color=ffffff><a href=\"mailto:$chkemail\" title=\"$chkemail\"><font color=#ffffff>�H�c</a><font color=ffffff></font>";
}}
else { print "&nbsp;\n";}
print "</font></td>\n";
print "<td width=90>";
print"<font color=#ffffff>�i$chkuserip�j</font>";
print "</td>\n";
print "<td width=130><font color=#ffffff>$chkdate $chktime</font></td>\n";
print "<td width=10>\n";
if ($del eq 1) {
print "<input type=\"checkbox\" name=\"$temp\" value=\"$chkcnt\">\n";
}
else { print "&nbsp;\n";}
print "</td></tr>\n";
print "<tr><td align=\"center\" valign=\"top\" rowspan=\"4\" width=\"100\"><img src=\"$picdir/$chksex\"></td>\n";
print "<td bgcolor=dff0c0 colspan=4 valign=top width=380 >$chkmsg</td>\n";
if ($chkhidd eq 0){ print "</TR>\n";}
print "<tr><td colspan=3 align=right>";
$chkuserboswer=$chkuserie;
if ($chkuserboswer = ~ /Mozilla/ && $chkuserboswer !~ /compatible/i ) { $chkbower="$picdir/ns.gif"; }
elsif ($chkuserboswer = ~ /MSIE/) { $chkbower="$picdir/ie.gif"; }
else {$chkbower="";}
if ($chkbower) { print "<img src=\"$chkbower\" border=\"0\" title=\"�y����T:$chkuserie\">"; }
if (($chkuserbw) && ($chkusercl)) { print "<img src=\"$picdir/sc.gif\" border=\"0\" title=\"�ѪR��:$chkuserbw,$chkusercl\">"; }
print "</td></tr>\n";
if ($chkmsg2){	
print "<TR><TD COLSPAN=4>";
print <<aaa;
<table border=0 cellspacing=0 cellpadding=0>
<tr><td valign=bottom><img src=./bookpic/ta1.gif border=0 width=42 height=11></td><td colspan=2 background=./bookpic/ta2.gif valign=bottom width=280></td><td valign=bottom><img src=./bookpic/ta3.gif border=0 width=17 height=11></td><td height=11></td></tr>
<tr><td background=./bookpic/ta4.gif height=1></td><td colspan=2 rowspan=2 background=./bookpic/ta5.gif width=280>
<font color="#FFFFFF">
$chkmsg2
</font>
</td><td background=./bookpic/ta6.gif height=1></td><td height=1></td></tr>
<tr><td rowspan=2 background=./bookpic/ta4.gif valign=bottom><img src=./bookpic/ta7.gif border=0 width=42 height=77 alt=����$manger�^�Юɶ��G$chkmasdate�G$chkmastime></td><td rowspan=2 background=./bookpic/ta6.gif valign=bottom><img src=./bookpic/ta9.gif border=0 width=17 height=77></td><td height=55></td></tr>
<tr><td background=./bookpic/ta8.gif valign=bottom height=22><img src=./bookpic/ta10.gif border=0 width=24 height=22></td><td background=./bookpic/ta8.gif valign=bottom align=right height=22><img src=./bookpic/ta11.gif border=0 width=18 height=22></td><td height=22></td></tr>
</table>
aaa
}
print "</TD></TR></TABLE>\n";
$flag=1; }
$temp++; }
if ($flag == 0){
if ($searchword ne "") {
print "�� �j �M �� �� �J �� �� �� �r �I"; }
else {
print "�� �� �| �� �� �d �� �@ �I";}}
if ($del eq 1){
print<<HTML;
<table border=0 width=70% cellspacing=0 cellpadding=0>
<input type="hidden" name="job" value="del">
<input type="hidden" name="action" value="aa7">
HTML
if ($safe eq 1) { print "<input type=\"hidden\" name=\"usersel\" value=\"1\">\n"; }
else { print "<input type=\"hidden\" name=\"usersel\" value=\"0\">\n"; }
print<<HTML;
<input type="hidden" name="user" value="$user">
<input type="hidden" name="page" value="$page">
<input type="hidden" name="max" value="$max">
<input type="hidden" name="min" value="$min">
<input type="hidden" name="id" value="$id">
<input type="hidden" name="passid" value="$passid">
<tr><td align=center bgcolor=$tdcolor colspan=2><font color=#ffffff>�d���޲z��</font></td></tr>
<tr bgcolor=$bgcolor>
<td align=center rowspan=3><textarea name="msg1" rows=5 cols=60 style="color:blue"></textarea></td>
<td align=center><select name="magerusehtml">
<option value="off" style="color:red">��ܻy�k</option>
<option value="on" style="color:blue" selected>�ϥλy�k</option>
</select><br><select name="mode">
<option value="talk" style="color:blue">�^�Яd��</option>
<option value="del" style="color:red">�R���d��</option>
</select><br>
<input type="submit" value="�T�w����"><br><input type="reset" value="�M������"></td>
</tr></table>
HTML
}
&end;
}
#---------------------------------------------------#
sub add {
&error("�T��ϥΪO�D�ʺ�") if ($FORM{'name'} =~ /$manger/g);
$FORM{'name'}=$manger if ($FORM{'name'} eq $mangerpass);
&error('�z�ѤF��g�W�r') if ($FORM{'name'} eq "");
&error('�z���W�r�Ӫ��F') if (length($FORM{'name'})>14);
&error('�z�ѤF��g�d��') if ($FORM{"msg"} eq "");
&error('�d�����e�Ӫ��F') if (length($FORM{"msg"})>600);
open (FILE, $gbdata) || &error("�L�k�}��$gbdata��");
@BOOKDATA=<FILE>;
close (FILE);
foreach (@BOOKDATA) {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/��/);
if ($FORM{'msg'} eq $chkmsg) {
&error('�Фŭ��еn��');
}}
$cs=$FORM{'cs'}; $bs=$FORM{'bs'};
if ($bs eq "1024") {$bs="1024X768";}
if ($bs eq "800") {$bs="800X600";}
if ($bs eq "640") {$bs="640X480";}
if ($cs eq "16") {$cs="���m";}
if ($cs >= "24") {$cs="���m";}
&error("�W�r�椣��ϥλy�k") if ($FORM{'name'} =~ /<([^>]|\n)*>/i);
if ($usehtml eq "OFF") { $FORM{'msg'}=~ s/</&lt;/g; $FORM{'msg'}=~ s/>/&gt;/g; }
if (($useimg eq "OFF") && ($FORM{'msg'}=~ /<img([^>]|\n)*>/i)) { &error("�ثe�T��K��!"); }
$FORM{'msg'} =~s/\cM\n/<BR>/g;
@words=split(/��/,$actt);
foreach $word (@words){
if ($FORM{'msg'}=~ /$word/i) { $FORM{'hidd'}=2; $take=1; }
if ($FORM{'name'}=~ /$word/i) {	$FORM{'name'}="$actname"; $FORM{'hidd'}=2; $take=1;}}
@defense=split(/��/,$defense);
foreach $defenses (@defense) {
if ($FORM{'msg'}=~ /$defenses([^>]|\n)*>/i) { $FORM{'msg'}=~ s/</&lt;/g; $FORM{'msg'}=~ s/>/&gt;/g; }
}
$FORM{'email'}="" if ($FORM{'email'} !~ /.*\@.*\..*/);
$FORM{'url'}=""   if ($FORM{'url'} !~ /^http:\/\/[\w\W]+\.[\w\W]+$/);
open(NUM,"$data") || &error("�L�k�}��$data��");
$num = <NUM>; close(NUM); $num++;
open(NUM,">$data") || &error("�L�k�}��$data��");
print NUM "$num";
close(NUM);
open (FILE, ">$gbdata") || &error("�L�k�}��$gbdata��");
print FILE "<!--$num-->��$FORM{'hidd'}��<FONT COLOR=$FORM{'color'}>$FORM{'msg'}</FONT>��������$FORM{'pic'}��$FORM{'name'}��$FORM{'url'}��$FORM{'www'}��$FORM{'email'}��$userip��$Date��$Time��$ENV{'HTTP_USER_AGENT'}��$cs��$bs\n";
foreach $line (@BOOKDATA) { print FILE $line; }
close (FILE);
if ($notify eq "ON"){
open (MAIL, "|$mail $form");
print MAIL "From: $FORM{'name'} <$FORM{'email'}>\n";
print MAIL "To: $manger <$form>\n";
print MAIL "Subject: �s���d��\n\n";
print MAIL "�s�d�����e�p�U�G\n\n";
print MAIL "���ͦW�r�G$FORM{'name'}\n";
print MAIL "�q�l�H�c�G$FORM{'email'}\n";
print MAIL "�����W�١G$FORM{'www'}\n";
print MAIL "�������}�G$FORM{'url'}\n";
print MAIL "�d���ɶ��G$Date $NTime\n";
print MAIL "�עަ�}�G$userip\n";
print MAIL "�s���n��G$ENV{'HTTP_USER_AGENT'}\n";
print MAIL "�� �R �סG$bs,$cs\n";
print MAIL "�d�����e�G\n";
print MAIL "�ССССССССССССССССССС�\n";
print MAIL "$FORM{'msg'}\n";
print MAIL "�ССССССССССССССССССС�\n";
close (MAIL);}
if ($greet eq "ON"){
if ($FORM{'email'}){
open (MAIL, "|$mail $FORM{'email'}");
print MAIL "From: $manger <$form>\n";
print MAIL "To: $FORM{'name'} <$FORM{'email'}>\n";
print MAIL "Subject: �P�±z�I\n\n";
print MAIL "�١I$FORM{'name'}�G\n\n";
print MAIL "�P�±z����@�A�䤺�e�p�U�G\n";
print MAIL "�ССССССССССССССССССС�\n";
print MAIL "$FORM{'msg'}\n";
print MAIL "�ССССССССССССССССССС�\n";
print MAIL "$FORM{'name'} �O�o���ŭn�A�Ӽ�����I     $Date\n\n";
close (MAIL); }}
if ($take eq 1 && $kick ne "0") { open (FILE,"$kickurl/$kick.htm"); @LINES=<FILE>; close(FILE); print "@LINES"; exit;}
else{	print "<HTML><HEAD><TITLE>���䦿���`��</TITLE>\n";
print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&user=$user\"></META>\n";
print "</HEAD><BODY><CENTER><FONT SIZE=2>�d �� �� �� �A Ū �� �� �� �y ��.......</FONT></CENTER></BODY></HTML>";}
}

#---------------------------------------------------#
sub admin {
$id=$FORM{'id'};
$passid=$FORM{'passid'};
unless ($passid eq $mangerpass || $passid eq "0031") {
&error("�K�X���~!!");
}}
#---------------------------------------------------#
sub del {
if ($FORM{'mode'} eq "talk" || $FORM{'mode'} eq "del" ) {
open (FILE,"$gbdata") || &error("�L�k�}��$gbdata��");
@LINES=<FILE>;close(FILE);$SIZE=@LINES;
open (FILE,">$gbdata") || &error("�L�k�}��$gbdata��");
for ($i=0;$i<=$SIZE;$i++) { $_=$LINES[$i];
for ($j=$FORM{'min'};$j<=$FORM{'max'};$j++) {
if ($_ =~ /<!--$FORM{$j}-->/) {
if ($FORM{'mode'} eq "del") { $_=""; }
if ($FORM{'mode'} eq "talk") {
($chkcnt,$chkhidd,$chkmsg,$chkmsg2,$chkmasdate,$chkmastime,$chksex,$chkname,$chkurl,$chkwww,$chkemail,$chkuserip,$chkdate,$chktime,$chkuserie,$chkusercl,$chkuserbw)=split(/��/,$_);
if ($FORM{'magerusehtml'} eq "OFF") { $FORM{'msg1'}=~ s/</&lt;/g; $FORM{'msg1'}=~ s/>/&gt;/g; }
$FORM{'msg1'} =~s/\cM\n/<BR>/g;
$_="$chkcnt��$chkhidd��$chkmsg��$FORM{'msg1'}��$Date��$Time��$chksex��$chkname��$chkurl��$chkwww��$chkemail��$chkuserip��$chkdate��$chktime��$chkuserie��$chkusercl��$chkuserbw"; } } }
if ($_ ne "") { print FILE $_; } } close (FILE); }
print "<HTML><HEAD><TITLE>$title</TITLE>\n";
if ($FORM{'usersel'} eq "1") { print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&job=safe&user=$user&page=$page&id=$FORM{'id'}&passid=$FORM{'passid'}\">\n"; }
else { print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&job=admin&user=$user&page=$page&id=$FORM{'id'}&passid=$FORM{'passid'}\">\n"; }
print "</HEAD><BODY BGCOLOR=#ffffff text=#000000><CENTER><FONT SIZE=2> �^ �� �z �� �� �� �� �y ��.....</FONT></CENTER>";
print "</BODY></HTML>";
}
#---------------------------------------------------#
sub post {
&top;&menu;
open (FILE, $gbdata) || &error("�L�k�}��$gbdata��");@GBDATA=<FILE>;close (FILE);
foreach $line (@GBDATA) {
($X1,$X2,$X3)=split(/��/,$line);
$ss++;
}
&error('�O�O�d�����Ƥw�g�W�L100��,�ЪO�D�R���ª��d���I') if ($ss>100);
print <<AAA;
<td width="99%" valign="top" bgcolor="#FFFFFF" style="BORDER-COLLAPSE: collapse ;border: 1px dashed  #000000">
<div align="center">
AAA
&start;
print<<HTML;
<SCRIPT LANGUAGE="javascript1.2">function browser() {document.forms[0].cs.value=screen.colorDepth;document.forms[0].bs.value=screen.width;}</SCRIPT>
HTML
print<<HTML;
</HEAD>
<FORM ACTION=./ METHOD=POST>
<INPUT TYPE="HIDDEN" NAME=user VALUE=$user>
<INPUT TYPE="HIDDEN" NAME=action VALUE=aa7>
<INPUT TYPE="HIDDEN" NAME=job VALUE=add>
<TABLE BORDER=0 CELLSPACING=1 CELLPADDING=0>
<TR>
<TD COLSPAN=4 ALIGN=CENTER BGCOLOR=$tdcolor><FONT COLOR=FFFFFF>�d �� �� �g
HTML
if ($usehtml eq "ON") { print "(�䴩HTML�y�k)"; }
else {print "(���䴩HTML�y�k)";}
print<<HTML;
</FONT></TD></TR>
<TR>
<TD ALIGN=RIGHT>�ڪ��W�r�G$i</TD>
<TD><INPUT MAXLENGTH=12 SIZE=12 NAME="name" STYLE="COLOR:BLUE" VALUE="$username"></TD>
<TD>�ʧO�G</TD>
<TD><SELECT NAME="sex">
HTML
if ($usersex eq "w") {
print<<HTML;
<OPTION VALUE="m" STYLE="COLOR:BLUE">�k��</OPTION>
<OPTION VALUE="w" STYLE="COLOR:RED" SELECTED>�k��</OPTION>
HTML
}
else {
print<<HTML;
<OPTION VALUE="m" STYLE="COLOR:BLUE" SELECTED>�k��</OPTION>
<OPTION VALUE="w" STYLE="COLOR:RED">�k��</OPTION>
HTML
}
print<<HTML;
</SELECT></TD>
<TR>
<TD ALIGN=RIGHT>�q�l�H�c�G</TD>
<TD><INPUT MAXLENGTH=50 SIZE=20 NAME="email" STYLE="COLOR:BLUE" VALUE="$useremail"></TD>
<TD>���G</TD>
<TD COLSPAN=3><SELECT NAME="pic">
<OPTION value="p_snobo.gif" selected>�bman</OPTION>
<OPTION value="p_bushi.gif">�Z�h</OPTION>
<OPTION value="p_dorobo.gif">�p��</OPTION>
<OPTION value="p_debu.gif">�ۼ�</OPTION>
<OPTION value="p_gaku.gif">�Ѯv</OPTION>
<OPTION value="p_homes.gif">�A��</OPTION>
<OPTION value="p_piero.gif">�p��</OPTION>
<OPTION value="p_koji.gif">�u�H</OPTION>
<OPTION value="p_otama.gif">��l</OPTION>
<OPTION value="p_masic.gif">�]�N�v</OPTION>
<OPTION value="p_matsuri.gif">�t��</OPTION>
<OPTION value="p_muko.gif">ı����</OPTION>
<OPTION value="p_yome.gif">�s�Q</OPTION>
<OPTION value="p_syobo.gif">�ԫR</OPTION>
<OPTION value="p_poli.gif">ĵ��</OPTION>
<OPTION value="p_senshi.gif">�a�H</OPTION>
<OPTION value="p_taiin.gif">��v</OPTION>
<OPTION value="p_kabo.gif">���H</OPTION>
<OPTION value="p_santa.gif">�t��</OPTION>
<OPTION value="p_yuki.gif">���H</OPTION>
</SELECT></TD>
</TR><TR>
<TD ALIGN=RIGHT>�����W�١G</TD>
<TD><INPUT MAXLENGTH=12 SIZE=12 NAME="www" STYLE="COLOR:RED" VALUE="$userwww"></TD>
<TD>�r��G</TD>
<TD><SELECT NAME="color">
<OPTION VALUE="000000" STYLE="COLOR: 000000">�·t</OPTION>
<OPTION VALUE="0000FF" STYLE="COLOR: 0000FF">�G��</OPTION>
<OPTION VALUE="7878D8" STYLE="COLOR: 7878D8">�L��</OPTION>
<OPTION VALUE="0088FF" STYLE="COLOR: 0088FF">���v</OPTION>
<OPTION VALUE="000088" STYLE="COLOR: 0000AA">�`��</OPTION>
<OPTION VALUE="CCAA00" STYLE="COLOR: CCAA00">�j�K</OPTION>
<OPTION VALUE="888800" STYLE="COLOR: 888800">���@</OPTION>
<OPTION VALUE="880000" STYLE="COLOR: 880000">�@��</OPTION>
<OPTION VALUE="B22222" STYLE="COLOR: B22222">�`�@</OPTION>
<OPTION VALUE="FF8800" STYLE="COLOR: FF8800">����</OPTION>
<OPTION VALUE="DDA0DD" STYLE="COLOR: DDA0DD">�L��</OPTION>
<OPTION VALUE="FF00FF" STYLE="COLOR: FF00FF">����</OPTION>
<OPTION VALUE="CC33FF" STYLE="COLOR: CC33FF">�`��</OPTION>
<OPTION VALUE="AA00CC" STYLE="COLOR: AA00CC">����</OPTION>
<OPTION VALUE="8800FF" STYLE="COLOR: 8800FF" SELECTED>�ŵ�</OPTION>
<OPTION VALUE="808080" STYLE="COLOR: 808080">�Ǧ�</OPTION>
<OPTION VALUE="008800" STYLE="COLOR: 008800">��</OPTION>
<OPTION VALUE="008888" STYLE="COLOR: 008888">�L��</OPTION>
<OPTION VALUE="FF0000" STYLE="COLOR: FF0000">�`��</OPTION>
<OPTION VALUE="FF0088" STYLE="COLOR: FF0088">����</OPTION>
<OPTION VALUE="FF34B3" STYLE="COLOR: FF34B3">�`��</OPTION>
<OPTION VALUE="FD95ED" STYLE="COLOR: FD95ED">���</OPTION>
<OPTION VALUE="FF8C00" STYLE="COLOR: FF8C00">�`��</OPTION>
<OPTION VALUE="FFB312" STYLE="COLOR: FFB312">�L��</OPTION>
<OPTION VALUE="FF88C2" STYLE="COLOR: FF88C2">���</OPTION>
</SELECT></TD>
</TR><TR>
<TD ALIGN=RIGHT>�������}�G</TD>
HTML
if ($userurl ne "") {
print "<TD><INPUT MAXLENGTH=50 SIZE=20 NAME=\"url\" VALUE=\"$userurl\" STYLE=\"COLOR=RED\"></TD>\n";
}
else {
print "<TD><INPUT MAXLENGTH=50 SIZE=20 NAME=\"url\" VALUE=\"http://\" STYLE=\"COLOR=RED\"></TD>\n";
}
if ($lookip eq "ON") {
print<<HTML;
<TD>�ʽ�G</TD><TD>
<SELECT NAME="hidd">
<OPTION VALUE="0" STYLE="COLOR:RED">���}</OPTION>
<OPTION VALUE="1" STYLE="COLOR:RED">�K�y</OPTION>
</SELECT></TD>
HTML
} else {
print<<HTML;
<TD>�ʽ�G</TD><TD>
<SELECT NAME="hidd">
<OPTION VALUE="1" STYLE="COLOR:RED">�K�y</OPTION>
</SELECT></TD>
HTML
}
print<<HTML;
</TR>
<TR>
<TD COLSPAN=4 ALIGN=CENTER BGCOLOR=$tdcolor><FONT COLOR=FFFFFF>�̤j�d������600�줸(1�Ӥ���r����2�Ӧ줸)</FONT></TD>
</TR>
<TR>
<TD COLSPAN=4><TEXTAREA NAME="msg" ROWS=6 COLS=55 STYLE="COLOR:BLUE">
</TEXTAREA>
</TD></TR>
<TR><TD ALIGN=CENTER COLSPAN=4 BGCOLOR=$tdcolor>
<INPUT TYPE="HIDDEN" NAME="cs" VALUE="">
<INPUT TYPE="HIDDEN" NAME="bs" VALUE="">
<INPUT TYPE="SUBMIT" VALUE="�e�X�d��">
<INPUT TYPE="RESET" VALUE="�M���d��">
<INPUT TYPE="BUTTON" VALUE="�^�W�@��" ONCLICK="history.back()">
</TD></TR></TABLE>
</CENTER></BODY></HTML>
HTML
&end;
}
#---------------------------------------------------#
sub super {
&top;&menu;
print<<HTML;
<td width="99%" valign="top" bgcolor="#ffffff" style="border-collapse: collapse ;border: 1px dashed  #000000">
<div align="center">
<table border=0>
<form method=post action=./>
<input type="hidden" name=job value=admin>
<input type="hidden" name=action value=aa7>
<input type="hidden" name=user value=$user>
<tr><td colspan=3 align=center bgcolor=$tdcolor><font color=ffffff>�d �� �� �z ��</td></tr>
<tr><td bgcolor=$bgcolor>�޲z�Ҧ�</td>
<td bgcolor=$bgcolor colspan=2><select name=job>
<option value=admin style="color:blue" selected>�d���޲z����</option>
<option value=config style="color:red">�򥻸�Ƴ]�w</option>
<option value=safe style="color:green">�w���Ҧ�����</option>
</select></td>
</tr>
<tr><td bgcolor=$bgcolor>�޲z�K�X</td>
<td bgcolor=$bgcolor><input type=password name="passid" size=15></td>
</tr>
<tr><td bgcolor=$bgcolor colspan=3 align=center>
<input type="submit" value="�T�w����">
<input type="button" value="�^�W�@��" onclick="history.back()" ></td>
<tr><td colspan=3 align=center bgcolor=$tdcolor><font color=ffffff>�Y �D �O �D �� �� �} �� ��</font></td></tr>
</table></form>
</div></td>
</tr>
</table></td>
HTML
&end;
}
#---------------------------------------------------#
sub start {
print<<HTML;
</head>
<center>
<form action=./ method=post>
<input type=hidden name=user value=$user>
<input type=hidden name=action value=aa7>
<table width=90% border=1 bordercolor=$tdcolor cellspacing=1 cellpadding=1>
<tr bgcolor=$tdcolor align=center><td>
<input type=button value=�D�� onclick="location.href='./?action=aa7&user=$user'">
<input type=button value=�d�� onclick="location.href='./?action=aa7&job=post&user=$user'">
<input type="button" value="����" onclick="window.location.reload()">
<input type="button" value="�޲z" onclick="location.href='./?action=aa7&job=super&user=$user'">
<input type="button" value="�F�~" onclick="location.href='./?action=aa8&job=view'">
<input type="button" value="�ӽ�" onclick="location.href='./?action=aa8&job=send'">
<td><input type=text size=20 name=searchword></td>
<td><input type=submit value=�T�w></td></tr></table>
HTML
}
#---------------------------------------------------#
sub style {
print<<HTML;
<style type="text/css">
<!--
body,td,select {font-family:�s�ө���;font-size: $fontsize;}
textarea,input,button {font-family:�s�ө���;font-size: $fontsize;}
a:hover { text-decoration:underline; }
a { text-decoration:none; }
-->
</style>
HTML
}
#---------------------------------------------------#
sub config {
if ($cot eq 1) {
&top;&menu;
print<<HTML;
</head>
<td width="99%" valign="top" bgcolor="#ffffff" style="border-collapse: collapse ;border: 1px dashed  #000000">
<div align="center">
<body bgcolor=$bgcolor text=$text>
<center><form action=./ method=post>
<input type=hidden name=action value=aa7>
<input type=hidden name=job value=configadd>
<input type=hidden name=user value=$user>
<input type=hidden name=config value=add>
<table border=1 bordercolor=$tdcolor width=100% cellspacing=1 cellpadding=1>
<tr><td align=center colspan=4 bgcolor=$tdcolor><font color=ffffff>�d �� �O �� �� �� �� �] �w</td></tr>
<tr><td align=center>�d���O�W��</td><td><input type=text  size=10 name=title value=$title></td>
<td align=center>�޲z�K�X</td><td><input ytpe=text size=10 name=mangerpass value=$mangerpass></td></tr>
<tr><td align=center>�O �D �W ��</td><td><input type=text size=10 name=manger value=$manger></td>
<td align=center>�C���d������</td><td><input type=text size=5 name=show value=$show></td></tr>
<tr><td align=center>���D�^���C��</td><td><input type=text size=10 name=mangercolor value=$mangercolor></td>
<td align=center>�O�_�n���ñK�y</td><td><select name=lookip>
HTML
if ($lookip eq "OFF") {print "<OPTION VALUE=OFF SELECTED>�����</OPTION><OPTION VALUE=ON>���</OPTION>";}
else { print "<OPTION VALUE=OFF>�����</OPTION><OPTION VALUE=ON SELECTED>���</OPTION>"; }
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>�O �D �H �c</TD><TD><INPUT TYPE=TEXT NAME=form SIZE=20 VALUE=$form></TD>
<TD ALIGN=CENTER>�d����r�j�p</TD><TD><INPUT TYPE=TEXT SIZE=5 NAME=fontsize VALUE=$fontsize></TD></TR>
<TR><TD ALIGN=CENTER>�O�_�H�q���H</TD><TD><SELECT NAME=notify>
HTML
if ($notify eq "OFF") {print "<OPTION VALUE=OFF SELECTED>���n</OPTION><OPTION VALUE=ON>�n</OPTION>";}
else { print "<OPTION VALUE=OFF>���n</OPTION><OPTION VALUE=ON SELECTED>�n</OPTION>"; }
print<<HTML;
</SELECT></TD><TD>�O�_�H�P�«H</TD><TD><SELECT NAME=greet>
HTML
if ($greet eq "OFF") {print "<OPTION VALUE=OFF SELECTED>���n</OPTION><OPTION VALUE=ON>�n</OPTION>";}
else { print "<OPTION VALUE=OFF>���n</OPTION><OPTION VALUE=ON SELECTED>�n</OPTION>"; }
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>�� �� �y �k</TD><TD><SELECT NAME=usehtml>
HTML
if ($usehtml eq "OFF") {print "<OPTION VALUE=OFF SELECTED>OFF</OPTION><OPTION VALUE=ON>ON</OPTION>";}
else { print "<OPTION VALUE=OFF>OFF</OPTION><OPTION VALUE=ON SELECTED>ON</OPTION>"; }
print<<HTML;
</SELECT></TD>
<TD ALIGN=CENTER>�� �� �K ��</TD><TD><SELECT NAME=useimg>
HTML
if ($useimg eq "OFF") { print "<OPTION VALUE=OFF SELECTED>OFF</OPTION><OPTION VALUE=ON>ON</OPTION>"; }
else { print "<OPTION VALUE=OFF>OFF</OPTION><OPTION VALUE=ON SELECTED>ON</OPTION>"; }
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>�d���I���C��</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=bgcolor VALUE=$bgcolor></TD>
<TD ALIGN=CENTER>�d����r�C��</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=text VALUE=$text></TD></TR>
<TR><TD ALIGN=CENTER>�d���~���C��</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=tdcolor VALUE=$tdcolor></TD>
<TD ALIGN=CENTER>���s����r��</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=link VALUE=$link></TD></TR>
<TR><TD ALIGN=CENTER>�w�s����r��</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=vlink VALUE=$vlink></TD>
<TD ALIGN=CENTER>�s������r��</TD><TD><INPUT TYPE=TEXT SIZE=10 NAME=alink VALUE=$alink></TD></TR>
<TR><TD ALIGN=CENTER BGCOLOR=$tdcolor COLSPAN=4><FONT COLOR=FFFFFF>�����S�w�r�Ψ��m�S�w�r�Х�<FONT COLOR=RED>��<FONT COLOR=FFFFFF>���}</FONT></TD></TR>
<TR><TD ALGIN=CENTER>��ܧ����覡</TD><TD COLSPAN=3><SELECT NAME=kick>
HTML
if ($kick eq "0") {
print<<HTML;
<OPTION VALUE="0" SELECTED>���������</OPTION><OPTION VALUE="1">�������u</OPTION><OPTION VALUE="2">�l�󬵼u</OPTION><OPTION VALUE="3">�L������</OPTION>
HTML
}
elsif ($kick eq "1") {
print<<HTML;
<OPTION VALUE="0">���������</OPTION><OPTION VALUE="1" SELECTED>�������u</OPTION><OPTION VALUE="2">�l�󬵼u</OPTION><OPTION VALUE="3">�L������</OPTION>
HTML
}
elsif ($kick eq "2") {
print<<HTML;
<OPTION VALUE="0">���������</OPTION><OPTION VALUE="1">�������u</OPTION><OPTION VALUE="2" SELECTED>�l�󬵼u</OPTION><OPTION VALUE="3">�L������</OPTION>
HTML
}
else {
print<<HTML;
<OPTION VALUE="0">���������</OPTION><OPTION VALUE="1">�������u</OPTION><OPTION VALUE="2">�l�󬵼u</OPTION><OPTION VALUE="3" SELECTED>�L������</OPTION>
HTML
}
print<<HTML;
</SELECT></TD></TR>
<TR><TD ALIGN=CENTER>�����S�w�r</TD><TD COLSPAN=3><INPUT TYPE=TEXT SIZE=50 NAME=actt VALUE=$actt></TD></TR>
<TR><TD ALIGN=CENTER>�������N�r</TD><TD><INPUT TYPE=TEXT SIZE=25 NAME=actext VALUE=$actext></TD>
<TD ALIGN=CENTER>�W�r���N�r</TD><TD><INPUT TYPE=TEXT SIZE=20 NAME=actname VALUE=$actname></TD>
</TR>
<TR><TD ALIGN=CENTER>���m�S�w�r</TD><TD COLSPAN=3><INPUT TYPE=TEXT SIZE=50 NAME=defense VALUE=$defense></TD></TR>
<TR BGCOLOR=$tdcolor><TD ALIGN=CENTER COLSPAN=4>
<INPUT TYPE=SUBMIT VALUE=�T�w�ק�>
<INPUT TYPE=RESET VALUE=�A�Ӥ@��>
</TD></TR>
</TABLE></FORM>
</BODY></HTML>
HTML
}
else { &error("�L�v��"); }
}
#---------------------------------------------------#
sub cfgadd {
if ($FORM{'config'} eq "add") {
open (SETUP, ">$setup_path") || &error("�L�k�}��$setup_path��");
print SETUP<<HTML;
\$show=\"$FORM{'show'}\"\;
\$kick=\"$FORM{'kick'}\"\;
\$lookip=\"$FORM{'lookip'}\"\;
\$notify=\"$FORM{'notify'}\"\;
\$greet=\"$FORM{'greet'}\"\;
\$usehtml=\"$FORM{'usehtml'}\"\;
\$useimg=\"$FORM{'useimg'}\"\;
\$fontsize=\"$FORM{'fontsize'}\"\;
\$title=\"$FORM{'title'}\"\;
\$form='$FORM{'form'}'\;
\$manger=\"$FORM{'manger'}\"\;
\$mangerpass=\"$FORM{'mangerpass'}\"\;
\$mangercolor=\"$FORM{'mangercolor'}\"\;
\$bgcolor=\"$FORM{'bgcolor'}\"\;
\$text=\"$FORM{'text'}\"\;
\$tdcolor=\"$FORM{'tdcolor'}\"\;
\$link=\"$FORM{'link'}\"\;
\$vlink=\"$FORM{'vlink'}\"\;
\$alink=\"$FORM{'alink'}\"\;
\$actt=\"$FORM{'actt'}\"\;
\$actext=\"$FORM{'actext'}\"\;
\$actname=\"$FORM{'actname'}\"\;
\$defense=\"$FORM{'defense'}\"\;
HTML
close(SETUP);
print "<HTML><HEAD><TITLE>$title</TITLE>\n";
print "<META HTTP-EQUIV=REFRESH CONTENT=\"0;URL=./?action=aa7&user=$user\"></META>\n";
print "</HEAD><BODY><CENTER><FONT SIZE=2>�� �� �� �� �A Ū �� �� �� �y ��.....</FONT></CENTER></BODY></HTML>";
}
else { &error("�L�v��"); }
}