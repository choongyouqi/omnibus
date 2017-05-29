<?php
header('Content-Type: text/plain');
date_default_timezone_set("Asia/Kuala_Lumpur");

//Month, Day, Year
$birthday = mktime(0,0,0, 8, 22, 1993);
$tenkday = strtotime("+10000 day", $birthday);
$now = time();

$labels = array(
	"Birthday" => $birthday, 
	"10k-Day" => $tenkday,
	"Today" => $now
);

foreach($labels as $label => $target)
{
	printDate($label, $target, $birthday);
}

function printDate($label, $time, $birthday)
{
	echo sprintf("%-10s\t%s\t%6s days\n", $label, date('Y-m-d H:i:s A T', $time), number_format(($time-$birthday)/86400));
}
?>