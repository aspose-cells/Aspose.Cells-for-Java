package com.aspose.cells.examples.articles;

public class WorkingWithPerl {

	my $ok = 0;BEGIN
	{ $| = 1; print "1..33\n"; }END
	{print "not ok $ok - is JavaServer on localhost running?\nJavaServer must be running for these tests to function.\n" unless $loaded;}

	BEGIN
	{
	print "WARNING: You cannot run these tests unless JavaServer is running!\n";
	print "Do you want to continue? (Y/n) ";
	my $in = <STDIN>;
	exit 1 if ($in =~ /^n/i);
	}
	use lib'.';
	use Java;
	my $java = new Java();$loaded=1;$ok++;print"ok $ok\n";
	my $workbook = $java -> create_object("com.aspose.cells.Workbook");$ok++;print"workbook $ok\n";#$workbook->

	open("t.xls");
	$ok++;
	print "open $ok\n";
	my $worksheets = $workbook->getWorksheets();
	$ok++;
	print "worksheets $ok\n";
	my $worksheet = $worksheets->get(0);
	$ok++;
	print "worksheet $ok\n";
	my $cells = $worksheet->getCells();
	$ok++;
	print "cells $ok\n";
	my $cell = $cells->getCell(0,1);
	$ok++;
	print "cell $ok\n";
	$cell->setValue(123);
	$cell = $cells->getCell(1,1);
	$cell->setValue(456);
	$cell = $cells->getCell(2,1);
	$cell->setFormula("=SUM(B1:B2)");
	$cell = $cells->getCell(3,1);
	$cell->setValue("abc");

	$workbook->save("t1.xls");
	$ok++;
	print "save $ok\n";

}
