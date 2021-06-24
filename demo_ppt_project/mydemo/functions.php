<?php

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

// ====================================================================
// create a new blank slide
// $newImgslide = newSlide($objPHPPowerPoint); use in this format

function newSlide(PhpOffice\PhpPresentation\PhpPresentation $objPHPPresentation){
    $slide = $objPHPPresentation->createSlide();
    return $slide;
}
// =====================================================================




// =====================================================================
// create a new blank slide with baground Image where baground image size must be in 16:9 ratio
// $currentSlide = createImgSlide('./rela2.jpg', '2nd slide', '2nd slide');   use in this format

function newImgslide($imgpath, $name, $description){
    // Create slide
    global $objPHPPowerPoint;
    $newImgslide = newSlide($objPHPPowerPoint);
    $shape = $newImgslide->createDrawingShape();
    $shape->setName($name)
      ->setDescription($description)
      ->setPath($imgpath)
      ->setHeight(775)
      ->setWidth(960)
      ->setOffsetX(0)
      ->setOffsetY(0);

return $newImgslide;
}
// =====================================================================




// =====================================================================
// function to add image to the slide
// imgpath, setname, setdesc as string always
function addImage($currentSlideName,$imgpath,$height,$width,$ofsetX,$ofsetY,$setname,$setdesc){
    $shape = $currentSlideName->createDrawingShape();
    $shape->setName($setname)
        ->setDescription($setdesc)
        ->setPath($imgpath)
        ->setHeight($height)
        ->setWidth($width)
        ->setOffsetX($ofsetX)
        ->setOffsetY($ofsetY);
}
// =====================================================================




// =====================================================================
// add slide heading to the current slide

function addSlideHeading($currentSlideName, $slideHeading,$fontStyle,$fontsize){
    $text = $currentSlideName->createRichTextShape()
      ->setHeight(50)
      ->setWidth(800)
      ->setOffsetX(70)
      ->setOffsetY(18);
$text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER );
$textRun = $text->createTextRun($slideHeading);
$textRun->getFont()->setBold(true)
                   ->setSize($fontsize)
                   ->setName($fontStyle)
                   ->getColor()->setRGB('000000');
}
// =====================================================================

function addNormalText($currentSlideName,$height,$width,$ofsetX,$ofsetY,$data,$fontStyle,$fontsize,$bold ){
    $text = $currentSlideName->createRichTextShape()
      ->setHeight($height)
      ->setWidth($width)
      ->setOffsetX($ofsetX)
      ->setOffsetY($ofsetY);
    $text->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_LEFT );
    $textRun = $text->createTextRun($data);
    $textRun->getFont()->setBold($bold)
                    ->setSize($fontsize)
                    ->setName('Calibri')
                    ->getColor()->setRGB('000000');
}



// =====================================================================
// add paragraph with bullet points
// e.g of block data
/* $block_data = array(
	"Appropriate training to all  doctors,nurses  and pharmacists  in regards to MAR writing, administration techniques, storage and dispensing respectively.",
	"Availability of formulary & E-Cims in all patient care areas through the intranet for clarifying drug related queries.",
    "Display of High Alert, LASA list throughout the hospital.",
    "Reported Medication errors presented in the PTC meeting and suggestions were made to respective department heads to take neccessary actions to minimize the medication errors."
); */
function addBulletText($currentSlideName,$height,$width,$ofsetX,$ofsetY,$block_data,$fontStyle,$fontsize,$bold ){
$shape = $currentSlideName->createRichTextShape()
      ->setHeight($height)
      ->setWidth($width)
      ->setOffsetX($ofsetX)
      ->setOffsetY($ofsetY);
$textRun = $shape->createTextRun(implode("\n\n", $block_data));
$textRun->getFont()->setBold($bold)
                   ->setSize($fontsize)
                   ->setName($fontStyle)
                   ->getColor()->setRGB('000000');
$oBulletStyle = $shape->getActiveParagraph()->getBulletStyle();
$oBulletStyle->setBulletType(PhpOffice\PhpPresentation\Style\Bullet::TYPE_BULLET);
$oBulletStyle->setBulletChar('• ');
}
// =====================================================================




// call this function like $shape = createTable(parameter list) 
// remember : here $shape will assign to the table $shape that was returned by the function
// =====================================================================
function createTable($currentSlideName,$height,$width,$ofsetX,$ofsetY,$numberOfColumn)
{
    $shape = $currentSlideName->createTableShape($numberOfColumn);
    $shape->setHeight($height);
    $shape->setWidth($width);
    $shape->setOffsetX($ofsetX);
    $shape->setOffsetY($ofsetY);

    return $shape;
}
// =====================================================================






/* here $shape which passed as a parameter is the table which was returned by the createtable function
this function makes the first row colored and other are white 

=============>>>> Table data must b like this i.e array of arrays
$table_data = array(
      array(data,data,data,data,data,data,data),
      array(data,data,data,data,data,data,data),
      array(data,data,data,data,data,data,data),
      array(data,data,data,data,data,data,data),
      array(data,data,data,data,data,data,data)
      

);*/

// =====================================================================
function createRow($shape, $table_data,$fontStyle){

    for ($row_no=0; $row_no< count($table_data);$row_no++) {
          $row_data = $table_data[$row_no];
          if ($row_no == 0) {
                # code...
                $row = $shape->createRow();
                $row->getFill()->setFillType(Fill::FILL_SOLID)
                   ->setStartColor(new Color('f54b42'))
                      ->setEndColor(new Color('f54b42'));
          } else {
                # code...
                if ($row_no % 2 == 0){

                    $row = $shape->createRow();
                    $row->getFill()->setFillType(Fill::FILL_SOLID)
                       ->setStartColor(new Color('dce0dd'))
                          ->setEndColor(new Color('dce0dd'));
                }
                else{
                    $row = $shape->createRow();
                    $row->getFill()->setFillType(Fill::FILL_SOLID)
                       ->setStartColor(new Color('9ca8d9'))
                          ->setEndColor(new Color('9ca8d9'));

                }
          }
          
        
    
            foreach ($row_data as $column) {
    
                $oCell = $row->nextCell();
                $oCell->getActiveParagraph()->getAlignment()->setHorizontal( Alignment::HORIZONTAL_CENTER )
                                                            ->setMarginLeft(5)->setMarginTop(10);
                                                          //   ->setMarginRight(5)->setMarginBottom(5);
                $textRun = $oCell->createTextRun($column);
                $textRun->getFont()->setBold(false)
                                   ->setSize(12)
                                   ->setName($fontStyle)
                                   ->getColor()->setRGB('000000');
                    };
        };
    
    }

// =====================================================================




/*e.g of seriesData
$seriesData = array('Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 80, 'mala'=>290);
*/

// =====================================================================
function addBarChat($currentSlideName,$height,$width,$ofsetX,$ofsetY,$seriesData,$title,$fontStyle){

global $oFill;
global $oShadow;

    // Create a bar chart (that should be inserted in a shape)
    $barChart = new Bar();
    $barChart->setGapWidthPercent(158);
    /////////// not not
    $series1 = new Series('', $seriesData);
    $series1->setShowSeriesName(false);
    $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('f54b42'));
    $series1->getFont()->getColor()->setRGB('FFFFFF');
    $barChart->addSeries($series1);

    $shape = $currentSlideName->createChartShape();
    $shape->setName('PHPPresentation Monthly Downloads')
        ->setResizeProportional(false)
        ->setHeight($height)
        ->setWidth($width)
        ->setOffsetX($ofsetX)
        ->setOffsetY($ofsetY);
    $shape->setFill($oFill);
    $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getTitle()->setText($title);
    $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
    $shape->getPlotArea()->getAxisX()->setTitle('');
    $shape->getPlotArea()->getAxisY()->getFont()->setName($fontStyle)->getColor()->setRGB('000000');
    $shape->getPlotArea()->getAxisY()->setTitle('');
    $shape->getPlotArea()->setType($barChart);
	$shape->getLegend()->setVisible(false);
}


// working on this function
// =====================================================================
function fnSlide_Pie3D($currentSlideName,$title,$seriesData)
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
    $shape->getTitle()->setText($title);
    $shape->getTitle()->getFont()->setItalic(true);
    $shape->getPlotArea()->setType($pie3DChart);
    $shape->getView3D()->setRotationX(30);
    $shape->getView3D()->setPerspective(30);
    $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
    $shape->getLegend()->getFont()->setItalic(true);
}
?>