<?php 
require_once("vendor/autoload.php"); 
use PhpOffice\PhpWord\Shared\Converter;
// Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();

/* Note: any element you append to a document must reside inside of a Section. */

// Adding an empty Section to the document...
$section = $phpWord->addSection();
// Adding Text element with font customized using explicitly created font style object...
$fontStyle = new \PhpOffice\PhpWord\Style\Font();
$fontStyle->setBold(true);
$fontStyle->setUnderline(true);
$fontStyle->setName('Times New Roman');
$fontStyle->setSize(20);
$fontStyle->setBgColor('yellow');
$myTextElement = $section->addText('PRESCRIPTION AUDIT SUMMARY REPORT');
$myTextElement->setFontStyle($fontStyle);
$section->addTextBreak(1);

$fontStyleName = 'oneUserDefinedStyle';
$phpWord->addFontStyle(
    $fontStyleName,
    array('name' => 'Times New Roman', 'size' => 12)
);
$section->addText(
    'Respected Dr. xxxxxx',
    $fontStyleName
);
$section->addText(
    'We, the pharmacology department, are conducting prescription audit on daily basis and we would like to draw your kind attention to certain noncompliance area in your prescription.',
    $fontStyleName
);
$section->addText(
    'TOTAL PRESCRIPTIONS AUDITED: 33',
    $fontStyleName
);

$styleTable = array('borderSize' => 6, 'borderColor' => '006699', 'cellMargin' => 20, 'cellPadding' => 50);
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
//charts
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
$section->addTextBreak(1);
$section->addText('Thanking You,',$fontStyleName,array("align" => "right"));
$section->addText('MGM PHARMACOLOGY DEPARTMENT',$fontStyleName,array("align" => "right"));

$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save(time().'_demo.docx');
print "Successfully Generated";
?>