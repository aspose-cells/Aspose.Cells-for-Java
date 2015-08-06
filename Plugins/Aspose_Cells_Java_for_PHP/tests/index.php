<?php

require_once("../../java/Java.inc");

require_once __DIR__ . '/../vendor/autoload.php'; // Autoload files using Composer autoload

use Aspose\Cells\WorkingWithFiles\FileHandlingFeatures\OpeningFiles;
use Aspose\Cells\WorkingWithFiles\FileHandlingFeatures\SavingFiles;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\ChartToImage;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\ConvertingToMhtmlFiles;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\Excel2PdfConversion;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\WorksheetToImage;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\ConvertingToXPS;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\ManagingDocumentProperties;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\ConvertingExcelFilesToHtml;
use Aspose\Cells\WorkingWithFiles\UtilityFeatures\ConvertingWorksheetToSVG;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\HideUnhideWorksheet;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\DisplayHideTabs;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\DisplayHideScrollBars;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\DisplayHideGridlines;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\PageBreakPreview;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\ZoomFactor;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\FreezePanes;
use Aspose\Cells\WorkingWithWorksheets\DisplayFeatures\SplitPanes;
use Aspose\Cells\WorkingWithWorksheets\ManagementFeatures\ManagingWorksheets\AddingWorksheetstoNewExcelFile;
use Aspose\Cells\WorkingWithWorksheets\ManagementFeatures\ManagingWorksheets\RemovingWorksheetsusingSheetName;
use Aspose\Cells\WorkingWithWorksheets\ManagementFeatures\ManagingWorksheets\RemovingWorksheetsusingSheetIndex;
use Aspose\Cells\WorkingWithWorksheets\SecurityFeatures\ProtectingWorksheet;
use Aspose\Cells\WorkingWithWorksheets\SecurityFeatures\UnprotectingPasswordProtectedWorksheet;
use Aspose\Cells\WorkingWithWorksheets\SecurityFeatures\UnprotectingSimplyProtectedWorksheet;
use Aspose\Cells\WorkingWithWorksheets\PageSetupFeatures\SettingPageOptions;
use Aspose\Cells\WorkingWithWorksheets\ValueFeatures\ManagingPageBreaks;
use Aspose\Cells\WorkingWithWorksheets\ValueFeatures\CopyingAndMovingWorksheets;
use Aspose\Cells\WorkingWithRowsAndColumns\RowsAndColumns;



print "Running Aspose\\Cells\\WorkingWithFiles\\FileHandlingFeatures\\OpeningFiles::run()" . PHP_EOL;
OpeningFiles::run(__DIR__ . '/data/WorkingWithFiles/FileHandlingFeatures/OpeningFiles/');

print "Running Aspose\\Cells\\WorkingWithFiles\\FileHandlingFeatures\\SavingFiles::run()" . PHP_EOL;
SavingFiles::run(__DIR__ . '/data/WorkingWithFiles/FileHandlingFeatures/SavingFiles/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\ChartToImage::run()" . PHP_EOL;
ChartToImage::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/ChartToImage/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\ConvertingToMhtmlFiles::run()" . PHP_EOL;
ConvertingToMhtmlFiles::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/ConvertingToMhtmlFiles/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\Excel2PdfConversion::run()" . PHP_EOL;
Excel2PdfConversion::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/Excel2PdfConversion/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\WorksheetToImage::run()" . PHP_EOL;
WorksheetToImage::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/WorksheetToImage/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\ConvertingToXPS::run()" . PHP_EOL;
ConvertingToXPS::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/ConvertingToXPS/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\ManagingDocumentProperties::run()" . PHP_EOL;
ManagingDocumentProperties::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/ManagingDocumentProperties/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\ConvertingExcelFilesToHtml::run()" . PHP_EOL;
ConvertingExcelFilesToHtml::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/ConvertingExcelFilesToHtml/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\ConvertingWorksheetToSVG::run()" . PHP_EOL;
ConvertingWorksheetToSVG::run(__DIR__ . '/data/WorkingWithFiles/UtilityFeatures/ConvertingWorksheetToSVG/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\HideUnhideWorksheet::run()" . PHP_EOL;
HideUnhideWorksheet::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/HideUnhideWorksheet/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\DisplayHideTabs::run()" . PHP_EOL;
DisplayHideTabs::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/DisplayHideTabs/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\DisplayHideScrollBars::run()" . PHP_EOL;
DisplayHideScrollBars::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/DisplayHideScrollBars/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\DisplayHideGridlines::run()" . PHP_EOL;
DisplayHideGridlines::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/DisplayHideGridlines/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\PageBreakPreview::run()" . PHP_EOL;
PageBreakPreview::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/PageBreakPreview/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\ZoomFactor::run()" . PHP_EOL;
ZoomFactor::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/ZoomFactor/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\FreezePanes::run()" . PHP_EOL;
FreezePanes::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/FreezePanes/');

print "Running Aspose\\Cells\\WorkingWithFiles\\UtilityFeatures\\SplitPanes::run()" . PHP_EOL;
SplitPanes::run(__DIR__ . '/data/WorkingWithWorksheets/DisplayFeatures/SplitPanes/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\ManagementFeatures\\ManagingWorksheets\\AddingWorksheetstoNewExcelFile::run()" . PHP_EOL;
AddingWorksheetstoNewExcelFile::run(__DIR__ . '/data/WorkingWithWorksheets/ManagementFeatures/ManagingWorksheets/AddingWorksheetstoNewExcelFile/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\ManagementFeatures\\ManagingWorksheets\\RemovingWorksheetsusingSheetName::run()" . PHP_EOL;
RemovingWorksheetsusingSheetName::run(__DIR__ . '/data/WorkingWithWorksheets/ManagementFeatures/ManagingWorksheets/RemovingWorksheetsusingSheetName/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\ManagementFeatures\\ManagingWorksheets\\RemovingWorksheetsusingSheetIndex::run()" . PHP_EOL;
RemovingWorksheetsusingSheetIndex::run(__DIR__ . '/data/WorkingWithWorksheets/ManagementFeatures/ManagingWorksheets/RemovingWorksheetsusingSheetIndex/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\SecurityFeatures\\ProtectingWorksheet::run()" . PHP_EOL;
ProtectingWorksheet::run(__DIR__ . '/data/WorkingWithWorksheets/SecurityFeatures/ProtectingWorksheet/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\SecurityFeatures\\UnprotectingPasswordProtectedWorksheet::run()" . PHP_EOL;
UnprotectingPasswordProtectedWorksheet::run(__DIR__ . '/data/WorkingWithWorksheets/SecurityFeatures/UnprotectingPasswordProtectedWorksheet/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\SecurityFeatures\\UnprotectingSimplyProtectedWorksheet::run()" . PHP_EOL;
UnprotectingSimplyProtectedWorksheet::run(__DIR__ . '/data/WorkingWithWorksheets/SecurityFeatures/UnprotectingSimplyProtectedWorksheet/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\PageSetupFeatures\\SettingPageOptions::run()" . PHP_EOL;
SettingPageOptions::run(__DIR__ . '/data/WorkingWithWorksheets/PageSetupFeatures/SettingPageOptions/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\ValueFeatures\\ManagingPageBreaks::run()" . PHP_EOL;
ManagingPageBreaks::run(__DIR__ . '/data/WorkingWithWorksheets/ValueFeatures/ManagingPageBreaks/');

print "Running Aspose\\Cells\\WorkingWithWorksheets\\ValueFeatures\\CopyingAndMovingWorksheets::run()" . PHP_EOL;
CopyingAndMovingWorksheets::run(__DIR__ . '/data/WorkingWithWorksheets/ValueFeatures/CopyingAndMovingWorksheets/');

print "Running Aspose\\Cells\\WorkingWithRowsAndColumns\\RowsAndColumns::run()" . PHP_EOL;
RowsAndColumns::run(__DIR__ . '/data/WorkingWithRowsAndColumns/RowsAndColumns/');
