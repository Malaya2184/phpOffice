<?php
//////// demo 1 start ///////////////////////////////


// with your own install
require_once 'src/PhpPresentation/Autoloader.php';
\PhpOffice\PhpPresentation\Autoloader::register();
require_once 'src/Common/Autoloader.php';
\PhpOffice\Common\Autoloader::register();
// require_once'./src/PhpPresentation/DocumentLayout.php';

// with Composer
// require_once 'vendor/autoload.php';
use PhpOffice\PhpPresentation\DocumentLayout;
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
$objPHPPowerPoint->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_SCREEN_16X9);

// $objPHPPowerPoint->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_CUSTOM, true)
// ->setCX( 1180,  DocumentLayout::UNIT_PIXEL)
// ->setCY( 768,  DocumentLayout::UNIT_PIXEL);

// for creation of a new slide
function createTemplatedSlide(PhpOffice\PhpPresentation\PhpPresentation $objPHPPresentation){
    // Create slide
    $slide = $objPHPPresentation->createSlide();
    return $slide;
}

// function to create rela slide using template image
function createRelaSlide($imgpath, $name, $description){
    // Create slide
    global $objPHPPowerPoint;
    $relaSlide = createTemplatedSlide($objPHPPowerPoint);
    $shape = $relaSlide->createDrawingShape();
    $shape->setName($name)
      ->setDescription($description)
      ->setPath($imgpath)
      ->setHeight(775)
      ->setWidth(960)
      ->setOffsetX(0)
      ->setOffsetY(0);

return $relaSlide;
}
function bulletText($currentSlideName,$height,$width,$ofsetX,$ofsetY,$block_data,$fontsize,$bold ){
      $shape = $currentSlideName->createRichTextShape()
            ->setHeight($height)
            ->setWidth($width)
            ->setOffsetX($ofsetX)
            ->setOffsetY($ofsetY);
      $textRun = $shape->createTextRun(implode("\n", $block_data));
      $textRun->getFont()->setBold($bold)
                         ->setSize($fontsize)
                         ->setName('Calibri (Body)')
                         ->getColor()->setRGB('000000');
      $oBulletStyle = $shape->getActiveParagraph()->getBulletStyle();
      $oBulletStyle->setBulletType(PhpOffice\PhpPresentation\Style\Bullet::TYPE_BULLET);
      $oBulletStyle->setBulletChar('• ');
      }


// Create slide 1
$currentSlide = $objPHPPowerPoint->getActiveSlide();

// Create a shape (drawing)
$shape = $currentSlide->createDrawingShape();
$shape->setName('PHPPresentation logo')
      ->setDescription('PHPPresentation logo')
      ->setPath('./rela123.jpg')
      ->setHeight(775)
      ->setWidth(960)
      ->setOffsetX(0)
      ->setOffsetY(0);

// $currentSlide = addImage($currentSlide ,'./resources/phppowerpoint_logo.gif', 40,80,5,5,'phplogo','phplog');


// Create a shape (text)
$text = $currentSlide->createRichTextShape()
      ->setHeight(120)
      ->setWidth(600)
      ->setOffsetX(60)
      ->setOffsetY(200);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $text->createTextRun('MEDICATION ERRORS');
$textRun->getFont()->setBold(false)
                   ->setSize(50)
                   ->setName('Calibri Light (Headings)')
                   ->getColor()->setRGB('000000');

//  slide 2
$currentSlide = createRelaSlide('./rela2.jpg', '2nd slide', '2nd slide');

// addImage($currentSlide ,'./resources/phppowerpoint_logo.gif', 40,80,5,5,'phplogo','phplog');
// slide 2 main heading
$text = $currentSlide->createRichTextShape()
      ->setHeight(50)
      ->setWidth(800)
      ->setOffsetX(70)
      ->setOffsetY(20);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $text->createTextRun('Medication error');
$textRun->getFont()->setBold(true)
                   ->setSize(22)
                   ->setName('Calibri Light (Headings)')
                   ->getColor()->setRGB('000000');
// addSlideHeading($currentSlide,'Medicational Error');
//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(65)
      ->setWidth(840)
      ->setOffsetX(70)
      ->setOffsetY(100);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Any error in the prescribing, transcribing, indenting, dispensing or administration of a drug, irrespective of whether such errors lead to adverse consequences or not, are the single most preventable cause of patient harm. A medication error is any preventable event that may cause or lead to inappropriate medication use or patient harm while the medication is in the control of the health care professional, patient, or consumer
');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//


//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(40)
      ->setWidth(840)
      ->setOffsetX(70)
      ->setOffsetY(210);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Medication errors represents the culture of patient safety in a hospital
To analyze the reasons of medication errors to identify the training needs & improvement opportunities to prevent the occurrence of same in future
');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(240)
      ->setOffsetX(70)
      ->setOffsetY(290);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Process & Outcome');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(400)
      ->setOffsetX(490)
      ->setOffsetY(290);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Incident forms, Appropriateness review & prescription audit form
');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(240)
      ->setOffsetX(70)
      ->setOffsetY(350);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('By 8th of every month');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(240)
      ->setOffsetX(490)
      ->setOffsetY(350);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('NA ');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(240)
      ->setOffsetX(70)
      ->setOffsetY(410);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Monthly');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(240)
      ->setOffsetX(490)
      ->setOffsetY(410);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Continuous');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(240)
      ->setOffsetX(70)
      ->setOffsetY(475);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Concurrent');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(240)
      ->setOffsetX(490)
      ->setOffsetY(475);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('All inpatient areas');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//


// slide 3
$currentSlide = createRelaSlide('./rela3.jpg', '3rd slide', '3rd slide');
$text = $currentSlide->createRichTextShape()
      ->setHeight(50)
      ->setWidth(800)
      ->setOffsetX(70)
      ->setOffsetY(20);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $text->createTextRun('Medication error');
$textRun->getFont()->setBold(true)
                   ->setSize(22)
                   ->setName('Calibri Light (Headings)')
                   ->getColor()->setRGB('000000');


//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(50)
      ->setWidth(250)
      ->setOffsetX(630)
      ->setOffsetY(240);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Errors in the prescribing, transcribing, dispensing, administering and monitoring of medications.');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//
//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(250)
      ->setOffsetX(630)
      ->setOffsetY(370);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('Error in OPD settings');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//
//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(25)
      ->setWidth(250)
      ->setOffsetX(630)
      ->setOffsetY(460);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun(' 7.68');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

// Create Table
// $shape = $currentSlide->createTableShape(7);
// $shape->setHeight(200);
// $shape->setWidth(700);
// $shape->setOffsetX(120);
// $shape->setOffsetY(450);

//Create Table
$shape = $currentSlide->createTableShape(7);
// $shape->setHeight(50);
// $shape->setWidth(550);
$shape->setOffsetX(90);
$shape->setOffsetY(300);

// Add row 1 ok-----------
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('ba610a'))
               ->setEndColor(new Color('ba610a'));
$oCell = $row->nextCell();
$oCell->setWidth(200);
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Marks ');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->setWidth(50);
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Oct’20');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->setWidth(50);
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Nov’20');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->setWidth(50);
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Dec’20');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->setWidth(50);
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Jan’21');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->setWidth(50);
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Feb’21');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->setWidth(50);
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Mar’21');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

// Add row 2 ok-----------
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('dce0dd'))
                        ->setEndColor(new Color('dce0dd'));
$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Total Number of Medicational Errors');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('11');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('0');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('4');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('7');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('6');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('5');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

// Add row 3 ok-----------
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('dce0dd'))
               ->setEndColor(new Color('dce0dd'));
$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Number of Patient Days');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('4243');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('2608');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('2722');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('2186');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('2122');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('2886');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

// Add row 4 ok-----------
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('dce0dd'))
               ->setEndColor(new Color('dce0dd'));
$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1000');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1000');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1000');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1000');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1000');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1000');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1000');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
// Add row 5 ok-----------
$row = $shape->createRow();
$row->getFill()->setFillType(Fill::FILL_SOLID)
			   ->setStartColor(new Color('dce0dd'))
               ->setEndColor(new Color('dce0dd'));
$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('Rate');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('2.59');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('0');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1.47');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('3.20');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('2.83');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

$oCell = $row->nextCell();
$oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $oCell->createTextRun('1.75');
$textRun->getFont()->setBold(false)
                   ->setSize(8)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');

// slide 4
$currentSlide = createRelaSlide('./rela4.jpg', '4th slide', '4th slide');
// addImage($currentSlide ,'./resources/phppowerpoint_logo.gif', 40,80,5,5,'phplogo','phplog');
$text = $currentSlide->createRichTextShape()
      ->setHeight(50)
      ->setWidth(800)
      ->setOffsetX(70)
      ->setOffsetY(20);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $text->createTextRun('Medication error');
$textRun->getFont()->setBold(true)
                   ->setSize(22)
                   ->setName('Calibri Light (Headings)')
                   ->getColor()->setRGB('000000');

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(80)
      ->setWidth(410)
      ->setOffsetX(70)
      ->setOffsetY(105);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('The data will be collected through Incident forms, Appropriateness review & prescription audit form. The analysis will be done by Quality & Clinical Pharmacist collectively and the same will be presented to management through Apex Quality improvement and  Pharmacy & Therapeutic Committee meetings.
');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//
$text = $currentSlide->createRichTextShape()
      ->setHeight(75)
      ->setWidth(400)
      ->setOffsetX(490)
      ->setOffsetY(110);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
$textRun = $text->createTextRun('The monthly compiled data will be circulated electronically by quality department to Clinical Pharmacist for further dissemination to all employees working in the department.
');
$textRun->getFont()->setBold(false)
                   ->setSize(10)
                   ->setName('Calibri')
                   ->getColor()->setRGB('000000');
//======================================================//

//=====================================================//

$shape = $currentSlide->createRichTextShape()
      ->setHeight(255)
      ->setWidth(410)
      ->setOffsetX(70)
      ->setOffsetY(240);
$block_data = array(
	"Timings of the prescription were written by nursing staff instead of  doctors.",
	"Elements of prescriptions not verified properly.",
    "Lack of availability of drug in markets/ pharmacy.",
    "Unawareness about the brand names of combination drugs.",
    "Illegibile prescription writing",
    " Use of unapproved abbrevations.",
    "Nurses didnot verify the 10 rights of administration."
);
$textRun = $shape->createTextRun(implode("\n", $block_data));
$textRun->getFont()->setBold(true)
                   ->setSize(14.5)
                   ->setName('Calibri (Body)')
                   ->getColor()->setRGB('000000');
$oBulletStyle = $shape->getActiveParagraph()->getBulletStyle();
$oBulletStyle->setBulletType(PhpOffice\PhpPresentation\Style\Bullet::TYPE_BULLET);
$oBulletStyle->setBulletChar('• ');

//====================================================//

//=====================================================//
$shape = $currentSlide->createRichTextShape()
      ->setHeight(295)
      ->setWidth(420)
      ->setOffsetX(490)
      ->setOffsetY(230);
$block_data = array(
	"Appropriate training to all  doctors,nurses  and pharmacists  in regards to MAR writing, administration techniques, storage and dispensing respectively.",
	"Availability of formulary & E-Cims in all patient care areas through the intranet for clarifying drug related queries.",
    "Display of High Alert, LASA list throughout the hospital.",
    "Reported Medication errors presented in the PTC meeting and suggestions were made to respective department heads to take neccessary actions to minimize the medication errors."
);
$textRun = $shape->createTextRun(implode("\n", $block_data));
$textRun->getFont()->setBold(true)
                   ->setSize(14.5)
                   ->setName('Calibri (Body)')
                   ->getColor()->setRGB('000000');
$oBulletStyle = $shape->getActiveParagraph()->getBulletStyle();
$oBulletStyle->setBulletType(PhpOffice\PhpPresentation\Style\Bullet::TYPE_BULLET);
$oBulletStyle->setBulletChar('• ');

//====================================================//

$currentSlide = createTemplatedSlide($objPHPPowerPoint);
global $oFill;
global $oShadow;

$series1Data = array('Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 80, 'mala'=>290);

    // Create a bar chart (that should be inserted in a shape)
    $barChart = new Bar();
    $barChart->setGapWidthPercent(158);
    ///////////
    $series1 = new Series('', $series1Data);
    $series1->setShowSeriesName(false);
    $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ba610a'));
    $series1->getFont()->getColor()->setRGB('FFFFFF');
    $barChart->addSeries($series1);

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



$currentSlide = createTemplatedSlide($objPHPPowerPoint);
function fnSlide_Pie3D($currentSlideName)
{
    global $oFill;
    global $oShadow;
    // Create templated slide
//     echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
    
    // Generate sample data for second chart
//     echo date('H:i:s') . ' Generate sample data for chart' . EOL;
    $seriesData = array('Monday' => 12, 'Tuesday' => 15, 'Wednesday' => 13, 'Thursday' => 17, 'Friday' => 14, 'Saturday' => 9, 'Sunday' => 7);
    // Create a pie chart (that should be inserted in a shape)
//     echo date('H:i:s') . ' Create a pie chart (that should be inserted in a chart shape)' . EOL;
    $pie3DChart = new Pie3D();
    $pie3DChart->setExplosion(20);
    $series = new Series(' ', $seriesData);
    $series->setShowSeriesName(true);
    $series->getDataPointFill(0)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4672A8'));
    $series->getDataPointFill(1)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFAB4744'));
    $series->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF8AA64F'));
    $series->getDataPointFill(3)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF725990'));
    $series->getDataPointFill(4)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4299B0'));
    $series->getDataPointFill(5)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFDC853E'));
    $series->getDataPointFill(6)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF93A9CE'));
    $pie3DChart->addSeries($series);
    // Create a shape (chart)
//     echo date('H:i:s') . ' Create a shape (chart)' . EOL;
    $shape = $currentSlideName->createChartShape();
    $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
//     $shape->setShadow($oShadow);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText('PHPPresentation Daily Downloads');
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getPlotArea()->setType($pie3DChart);
    $shape->getView3D()->setRotationX(30);
    $shape->getView3D()->setPerspective(30);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}
fnSlide_Pie3D($currentSlide);

$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
print "created Successfully";




?>