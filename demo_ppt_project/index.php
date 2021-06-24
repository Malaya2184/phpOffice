<?php
// requirements
require './mydemo/functions.php';
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

$fontStyle1= 'Roboto Condensed';
$fontStyle2= 'Bahnschrift SemiBold';


$objPHPPowerPoint = new PhpPresentation();
// document layout set to 16:9 otherwise functions will not work perfectly
$objPHPPowerPoint->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_SCREEN_16X9);

// slide 1
$currentSlide = $objPHPPowerPoint->getActiveSlide();
addImage($currentSlide,'./img/light_grey.jpg',775,960,0,0,'first slide','first slide');
addImage($currentSlide,'./img/logo.png',100,240,350,30,'logo','logo');
addNormalText($currentSlide,250,450,450,250,'THIS IS PRESENTATION TITLE',$fontStyle1,45,true );
addImage($currentSlide,'./img/HEAD.png',150,190,230,250,'HEAD','HEAD');

// slide 2

$currentSlide = newImgslide('./img/light_grey.jpg', 'second slide', 'second slide');
addImage($currentSlide,'./img/logo.png',50,120,5,5,'logo','logo');
addSlideHeading($currentSlide, 'This Is Slide Heading',$fontStyle1,22);
addImage($currentSlide,'./img/orng.png',400,450,30,70,'orng','orng');
addNormalText($currentSlide,40,450,30,70,'This Is Sub Heading 1',$fontStyle1,19,true );
$block_data= array(
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
);
addBulletText($currentSlide,550,450,30,120,$block_data,$fontStyle2,14,true );
addImage($currentSlide,'./img/grn.png',400,450,480,70,'grn','grn');
addNormalText($currentSlide,40,450,480,70,'This Is Sub Heading 2',$fontStyle1,19,true );
$block_data= array(
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
    'Lorem Ipsum is simply dummy text of the printing and typesetting industry.',
);
addBulletText($currentSlide,550,450,480,120,$block_data,$fontStyle2,14,true );



// slide 3
$currentSlide = newImgslide('./img/light_grey.jpg', 'second slide', 'second slide');
addImage($currentSlide,'./img/logo.png',50,120,5,5,'logo','logo');
addSlideHeading($currentSlide, 'This Is Slide Heading',$fontStyle1,22);
addNormalText($currentSlide,40,450,30,70,'This Is Sub Heading 1',$fontStyle1,19,true );
$data = "Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. 

It has survived not only five centuries, but also the leap into electronic typesetting, remaining essentially unchanged. It was popularised in the 1960s with the release of Letraset sheets containing Lorem Ipsum passages, and more recently with desktop publishing software like Aldus PageMaker including versions of Lorem Ipsum

Lorem Ipsum is simply dummy text of the printing and typesetting industry. Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when an unknown printer took a galley of type and scrambled it to make a type specimen book. It has survived not only five centuries, but also the leap into electronic typesetting,";
addNormalText($currentSlide,380,900,30,120,$data,$fontStyle1,17,false );

// slide 4
$currentSlide = newImgslide('./img/light_grey.jpg', 'second slide', 'second slide');
addImage($currentSlide,'./img/logo.png',50,120,5,5,'logo','logo');
addSlideHeading($currentSlide, 'This Is Slide Heading',$fontStyle1,22);
$data = "Lorem ipsum dolor sit amet consectetur adipisicing elit. Tenetur, dolores repellendus! Magnam repellendus odit quo expedita cupiditate! Quaerat, voluptatibus doloremque.Lorem ipsum dolor sit amet consectetur adipisicing elit. Tenetur, dolores repellendus! Magnam repellendus odit quo expedita cupiditate! Quaerat, voluptatibus doloremque." ;
addNormalText($currentSlide,40,900,30,70,'This Is Sub Heading 1',$fontStyle1,19,true );
addNormalText($currentSlide,150,900,30,120,$data,$fontStyle1,17,false );
addNormalText($currentSlide,40,900,30,275,'This Is Sub Heading 2',$fontStyle1,19,true );
addNormalText($currentSlide,150,900,30,325,$data,$fontStyle1,17,false );


// slide 5
$currentSlide = newImgslide('./img/light_grey.jpg', 'second slide', 'second slide');
addImage($currentSlide,'./img/logo.png',50,120,5,5,'logo','logo');
addSlideHeading($currentSlide, 'This Is Slide Heading',$fontStyle1,22);
$seriesData = array('ABC' => 200, 'DEF' => 250, 'GHI' => 400, 'IJK' => 300, 'LMN' => 180, 'OPQ' => 80, 'RST'=>290,'UVW'=>360, 'XYZ'=>420);
addBarChat($currentSlide,300,900,35,60,$seriesData,'My Bar Chat',$fontStyle1);
$shape = createTable($currentSlide,180,900,35,370,10);
$table_data = array(
    array('index','index','index','index','index','index','index','index','index','index'),
    array('data','data','data','data','data','data','data','data','data','data'),
    array('data','data','data','data','data','data','data','data','data','data'),
    array('data','data','data','data','data','data','data','data','data','data'),

);
createRow($shape, $table_data,$fontStyle1);

// slide 6
$currentSlide = newImgslide('./img/light_grey.jpg', 'second slide', 'second slide');
addImage($currentSlide,'./img/thanks.png',250,500,220,200,'logo','logo');



$oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
$oWriterPPTX->save(__DIR__ . "/sample.pptx");
print "created Successfully";

?>