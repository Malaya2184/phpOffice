<?php
require_once("vendor/autoload.php"); 
use PhpOffice\PhpWord\Shared\Converter;
use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\SimpleType\VerticalJc;
use PhpOffice\PhpWord\Element\Chart;
use \PhpOffice\PhpWord\TemplateProcessor;


// // Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();
//==================================================
$time = date("h-i-sa",time());
$fontStyle2 = new \PhpOffice\PhpWord\Style\Font();
$fontStyle2->setBold(true);
$fontStyle2->setUnderline(false);
$fontStyle2->setName('Roboto Condensed');
$fontStyle2->setSize(20);
$fontStyle2->setColor('1e223b');



// 'borderColor' => '0d0d0d', 'borderSize' => 12, 
// page 1
$section =$phpWord->addSection(array('borderColor' => '00FF00', 'borderSize' => 12));
$section->addImage(
    './img/logo.png',
    array(
        'width'            => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(2.8),
        'height'           => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(0.9),
        'marginTop'     => \PhpOffice\PhpWord\Shared\Converter::cmToPixel(-3.55),
        'alignment'        => \PhpOffice\PhpWord\SimpleType\Jc::CENTER

    )
);
$phpWord->addParagraphStyle('pstyle', array('align'=>'center'));
$fontStyle = new \PhpOffice\PhpWord\Style\Font();
$fontStyle->setBold(true);
$fontStyle->setUnderline(false);
$fontStyle->setName('Roboto Condensed');
$fontStyle->setSize(40);
$fontStyle->setColor('1e223b');
$myTextElement = $section->addText('PRESCRIPTION AUDIT SUMMARY REPORT', $fontStyle, 'pstyle');

$section->addTextBreak(20);
$myTextElement = $section->addText("Time of generation : $time ", $fontStyle2, 'pstyle');



/* Note: any element you append to a document must reside inside of a Section. */
// Adding an empty Section to the document...
// $section = $phpWord->addSection(array('borderColor' => '0d0d0d', 'borderSize' => 12));
$section = $phpWord->addSection(array('borderColor' => '00FF00', 'borderSize' => 12));
// Adding Text element with font customized using explicitly created font style object...
$fontStyle = new \PhpOffice\PhpWord\Style\Font();
$fontStyle->setBold(true);
$fontStyle->setUnderline(true);
$fontStyle->setName('Times New Roman');
$fontStyle->setSize(20);
$fontStyle->setBgColor('yellow');
$myTextElement = $section->addText('PRESCRIPTION AUDIT SUMMARY REPORT',$fontStyle, array('align' => 'left'));

$section->addTextBreak(1);

$phpWord->addParagraphStyle('p1style', array('align'=>'both'));
$fontStyleName = 'oneUserDefinedStyle';
$phpWord->addFontStyle(
    $fontStyleName,
    array('name' => 'Times New Roman', 'size' => 12, 'underline' => 'single', 'align'=>'both')
);
$section->addText(
    'Respected Dr. xxxxxx',
    $fontStyleName,'p1style'
);
$section->addText(
    'We, the pharmacology department, are conducting prescription audit on daily basis and we would like to draw your kind attention to certain noncompliance area in your prescription.',
    $fontStyleName,'p1style'
);
$section->addText(
    'TOTAL PRESCRIPTIONS AUDITED: 33',
    $fontStyleName,'p1style'
);


// =====================================================

$styleTable = array('borderSize' => 6, 'borderColor' => '006699', 'cellMargin' => 20, 'cellPadding' => 50,'unit' => \PhpOffice\PhpWord\Style\Table::WIDTH_PERCENT, 'width' => 100 * 50);
$styleFirstRow = array('borderBottomSize' => 18, 'borderBottomColor' => '0000FF', 'bgColor' => '66BBFF');
$styleCell = array('valign' => 'center');
$fontStyle = array('bold' => true, 'align' => 'center');
$phpWord->addTableStyle('Fancy Table', $styleTable, $styleFirstRow);
$table = $section->addTable('Fancy Table');
$table->addRow();
$table->addCell(800, $styleCell)->addText(htmlspecialchars('Row 1'), $fontStyle);
$table->addCell(5000, $styleCell)->addText(htmlspecialchars('Row 2'), $fontStyle);
$table->addCell(800, $styleCell)->addText(htmlspecialchars('Row 3'), $fontStyle);
$table->addCell(800, $styleCell)->addText(htmlspecialchars('Row 4'), $fontStyle);
for ($i = 1; $i <= 8; $i++) {
    $table->addRow();
    $table->addCell(800)->addText(htmlspecialchars("Cell {$i}"));
    $table->addCell(5000)->addText(htmlspecialchars("Cell {$i}"));
    $table->addCell(800)->addText(htmlspecialchars("Cell {$i}"));
    $table->addCell(800)->addText(htmlspecialchars("Cell {$i}"));
}
$section->addTextBreak(1);

// //===================================================

// //charts
$section = $phpWord->addSection(array('colsNum' => 1, 'breakType' => 'continuous'));
$chartTypes = 'column';
$categories = array('N1', 'N2', 'N3', 'N4', 'N5', 'N6', 'N7', 'N6', 'N7', 'N8', 'N9', 'N10', 'N11', 'N12', 'N13','N14');
$series1 = array(1, 3, 2, 5, 4,5,6,7,8,9,10,11,12,13,14,15);
$showGridLines = false;
$showAxisLabels = true;
$styleChart = array('dataLabelOptions' => array('showCatName' => false,'showVal' => true));
$chart = $section->addChart($chartTypes, $categories, $series1,$styleChart);
$chart->getStyle()->setWidth(Converter::inchToEmu(5.5))->setHeight(Converter::inchToEmu(3));
$chart->getStyle()->setShowGridX($showGridLines);
$chart->getStyle()->setShowGridY($showGridLines);
$chart->getStyle()->setShowAxisLabels($showAxisLabels);

$section->addText(
    'As the compliance to prescription parts can greatly avoid near miss errors and medication
errors or any other adverse drug events, we humbly request your immense support to improve
the patient care service.',
    $fontStyleName
);
$section->addText('Thanking You,',$fontStyleName,array("align" => "right"));
$section->addText('MGM PHARMACOLOGY DEPARTMENT',$fontStyleName,array("align" => "right"));


// ==================================================================================================================

// word file creation
date_default_timezone_set('Asia/Calcutta');
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save(date("H-i-sa",time()).'_demo.docx');
print "Successfully Generated";
?>