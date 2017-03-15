use strict;
use Spreadsheet::WriteExcel;

my $xname = 'People_02.xml';
my $ename = 'Pacients.xls';
my $excel;
my $list1;
my $style;
my $tstyle;

my $title = " ";
my $column = " ";
my $field;
my $i = 1;
my $c;

my @db;

$/ = ">";
#xml file ------------------------------------
open xml, $xname or die "Couldn't open: $!\n";

#excel file ----------------------------------
$excel = Spreadsheet::WriteExcel->new($ename);
$list1 = $excel->add_worksheet('pacients');

$style = $excel->add_format(align=>'right');
$style->set_bg_color('silver');
$style->set_border();

$tstyle = $excel->add_format(align=>'right');
$tstyle->set_bg_color('cyan');
$tstyle->set_bold();
$tstyle->set_border();

$list1->write("A1", "Age", $tstyle);
$list1->write("B1", "Amount", $tstyle);

while ($field = <xml>)
{
	if ($field =~ /agecaption="(\d+)/)
	{
		$column = $1;
		
		if ($field =~ /agecaption="(\d+-\d+)/)
		{
			$column = "$1 m";
		}
	}
	
	if($column ne $title)
	{
		$title = $column;
		$i++;
		@db[$i] = 0;
		$list1->write("A$i", $title, $style);
		
		seek (xml, 0, 1);
		print "Been read bytes..........";
		print tell();
		print "\n";
	}

	if ($field =~ /cntpeople="(\d+)"/)
	{
		@db[$i] += $1;
	}	
}

for($c = 2; $c <= $i; $c++)
{
	$list1->write("B$c", @db[$c], $style);
}

print "Done.....................ok";
exit;