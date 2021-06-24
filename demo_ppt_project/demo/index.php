<?php
// with your own install
require_once 'src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();
require_once 'src/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();

// with Composer
//require_once 'vendor/autoload.php';

use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Pie3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Shadow;

$objPHPPowerPoint = new PhpPresentation();

function fnSlide_Bar(PhpPresentation $objPHPPresentation) {
    global $oFill;
    global $oShadow;

    // Create templated slide
    $currentSlide = createTemplatedSlide($objPHPPresentation);

    // Generate sample data for first chart
    $series1Data = array('Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 80);

    // Create a bar chart (that should be inserted in a shape)
    $barChart = new Bar();
    $barChart->setGapWidthPercent(158);
    /////////// not not
    $series1 = new Series('', $series1Data);
    $series1->setShowSeriesName(false);
    $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ba610a'));
    $series1->getFont()->getColor()->setRGB('FFFFFF');
    $barChart->addSeries($series1);

    // Create a shape (chart)
    $shape = $currentSlide->createChartShape();
    $shape->setName('PHPPresentation Monthly Downloads')
        ->setResizeProportional(false)
        ->setHeight(350)
        ->setWidth(700)
        ->setOffsetX(120)
        ->setOffsetY(80);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('DATA ANALYSIS');
    $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    $shape->getPlotArea()->getAxisX()->setTitle('');
    $shape->getPlotArea()->getAxisY()->getFont()->getColor()->setRGB('000000');
    $shape->getPlotArea()->getAxisY()->setTitle('');
    $shape->getPlotArea()->setType($barChart);
	$shape->getLegend()->setVisible(false);
	return $currentSlide;
}

function createTemplatedSlide(PhpOffice\PhpPresentation\PhpPresentation $objPHPPresentation){
    // Create slide
    $slide = $objPHPPresentation->createSlide();
    return $slide;
}


// ok----------------
// Create slide
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Create a shape (drawing) 
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPresentation logo')
      ->setDescription('PHPPresentation logo')
      ->setPath('./resources/phppowerpoint_logo.gif')
      ->setHeight(36)
      ->setOffsetX(10)
      ->setOffsetY(10);
$shape->getShadow()->setVisible(true)
                   ->setDirection(45)
                   ->setDistance(10);

// Create a shape (text)
$shape = $currentSlide->createRichTextShape()
      ->setHeight(300)
      ->setWidth(600)
      ->setOffsetX(170)
      ->setOffsetY(180);
$shape->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $shape->createTextRun('Univate Hospital');
$textRun->getFont()->setBold(true)
                   ->setSize(40)
                   ->setColor( new Color( 'FFE06B20' ) );
$textRun = $shape->getActiveParagraph()->createBreak();
$textRun = $shape->createTextRun('Department - Pharmacy');
$textRun->getFont()->setBold(true)
                   ->setSize(40)
                   ->setColor( new Color( 'FFE06B20' ) );
				   
// ok----------------

// not not -------------------
$oShadow = new Shadow();
$oShadow->setVisible(true)->setDirection(45)->setDistance(10);
// not not -------------------

$currentSlide = fnSlide_Bar($objPHPPowerPoint);

//Create Table
$shape = $currentSlide->createTableShape(7);
$shape->setHeight(200);
$shape->setWidth(700);
$shape->setOffsetX(120);
$shape->setOffsetY(450);



// Add row ok-----------
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('ba610a'))
               ->setEndColor(new Color('ba610a'));
$oCell = $row->nextCell();
$oCell->createTextRun(' ');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('Jul');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('Aug');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('Sep');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('Oct');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('Nov');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('Dec');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);


// Add row
$row = $shape->createRow();
$oCell = $row->nextCell();
$oCell->createTextRun('Numerator');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R3C2');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R3C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);

// Add row
$row = $shape->createRow();
$oCell = $row->nextCell();
$oCell->createTextRun('Denominator');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R3C2');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R3C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);

// Add row
$row = $shape->createRow();
$oCell = $row->nextCell();
$oCell->createTextRun('KPI Value');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R3C2');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R3C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);
$oCell = $row->nextCell();
$oCell->createTextRun('R2C3');
$oCell->getActiveParagraph()->getAlignment()->setMarginLeft(10)->setMarginTop(10);

// ok-----------

//Paragraph with bullets
$currentSlide = createTemplatedSlide($objPHPPowerPoint);
$shape = $currentSlide->createRichTextShape()
      ->setHeight(50)
      ->setWidth(800)
      ->setOffsetX(100)
      ->setOffsetY(100);
$textRun = $shape->createTextRun('DATA AGGREGATION AND DISSEMINATION');
$textRun->getFont()->setBold(true)->setSize(20);
$shape = $currentSlide->createRichTextShape()
      ->setHeight(300)
      ->setWidth(800)
      ->setOffsetX(100)
      ->setOffsetY(150);
$block_data = array(
	"Data aggregation Plan :The data will be collected from emergency medicine checklist. The analysis will be done by chief pharmacist & Quality Department collectively. The collected data will be presented in Pharmacy therapeutic committee, and the same will be presented to management through Quality  Patient safety committee meeting.",
	"Data Dissemination Plan : The monthly compiled data will be communicated  to all employees working in the department for further dissemination by Department in charge.",
);
$textRun = $shape->createTextRun(implode("\n", $block_data));
$textRun->getFont()->setSize(18);
$oBulletStyle = $shape->getActiveParagraph()->getBulletStyle();
$oBulletStyle->setBulletType(PhpOffice\PhpPresentation\Style\Bullet::TYPE_BULLET);
$oBulletStyle->setBulletChar('• ');


$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
print "created Successfully";
//$oWriterODP = IOFactory::createWriter($objPHPPowerPoint, 'ODPresentation');
//$oWriterODP->save(__DIR__ . "/sample.odp");
?>