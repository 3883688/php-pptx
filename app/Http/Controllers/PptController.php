<?php

namespace App\Http\Controllers;

use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Http\Request;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\IOFactory;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\AutoShape;
use PhpOffice\PhpPresentation\Shape\Chart\Axis;
use PhpOffice\PhpPresentation\Shape\Chart\Gridlines;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Area;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar3D;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Scatter;
use PhpOffice\PhpPresentation\Slide\SlideLayout;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;


use PhpOffice\PhpPresentation\Shape\Chart;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Outline;
use PhpOffice\PhpPresentation\Style\Shadow;
use PhpOffice\PhpSpreadsheet\Chart\DataSeries;

define('EOL', PHP_EOL);

class PptController extends Controller {
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

    public function index() {
        $objPHPPresentation = new PhpPresentation();
        // Set properties
        echo date('H:i:s') . ' Set properties' . PHP_EOL;
        $objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')->setTitle('Sample 07 Title')
            ->setSubject('Sample 07 Subject')->setDescription('Sample 07 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // Remove first slide
        echo date('H:i:s') . ' Remove first slide' . PHP_EOL;
        $objPHPPresentation->removeSlideByIndex(0);

        // Set Style
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));

        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

        // Generate sample data for chart
        echo date('H:i:s') . ' Generate sample data for chart' . PHP_EOL;
        $seriesData = [
            'Monday 01'    => 12,
            'Tuesday 02'   => 15,
            'Wednesday 03' => 13,
            'Thursday 04'  => 17,
            'Friday 05'    => 14,
            'Saturday 06'  => 9,
            'Sunday 07'    => 7,
            'Monday 08'    => 8,
            'Tuesday 09'   => 8,
            'Wednesday 10' => 15,
            'Thursday 11'  => 16,
            'Friday 12'    => 14,
            'Saturday 13'  => 14,
            'Sunday 14'    => 13,
        ];

        // Create templated slide
        echo PHP_EOL . date('H:i:s') . ' Create templated slide' . PHP_EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

        // Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . PHP_EOL;
        $lineChart = new Line();
        $series    = new Series('Downloads', $seriesData);
        $series->setShowSeriesName(true);
        $series->setShowValue(true);
        $lineChart->addSeries($series);

        // Create a shape (chart)
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit(3);
        $shape->getPlotArea()->getAxisY()->setMajorUnit(5);


        // Create a line chart (that should be inserted in a shape)
        $oOutline = new Outline();
        // 散点图线条类型
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_DARKRED));
        $oOutline->setWidth(2);

        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . PHP_EOL;
        $series = $lineChart->getSeries();
        $series[0]->setOutline($oOutline);
        $series[0]->getMarker()->setSymbol(Marker::SYMBOL_DIAMOND);
        $series[0]->getMarker()->setSize(7);
        $lineChart->setSeries($series);

        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function xyE() {
        $objPHPPresentation = new PhpPresentation();

        /*
        $seriesData = [
            'Monday'    => 0.1,
            'Tuesday'   => 0.33333,
            'Wednesday' => 0.4444,
            'Thursday'  => 0.5,
            'Friday'    => 0.4666,
            'Saturday'  => 0.3666,
            'Sunday'    => 0.1666
        ];

        // Create a scatter chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a scatter chart (that should be inserted in a chart shape)' . EOL;
        $lineChart = new Scatter();
        $lineChart->setIsSmooth(true);
        $series = new Series('Downloads', $seriesData);
        $series->setShowSeriesName(true);
        $series->getMarker()->setSymbol(Marker::SYMBOL_CIRCLE);
        $series->getMarker()->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->setStartColor(new Color('FF6F3510'))
            ->setEndColor(new Color('FF6F3510'));
        $series->getMarker()->getBorder()->getColor()->setRGB('FF0000');
        $series->getMarker()->setSize(10);
        $lineChart->addSeries($series);
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);
        // Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart)' . EOL;
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Daily Download Distribution')
            ->setResizeProportional(false)
            ->setHeight(550)
            ->setWidth(700)
            ->setOffsetX(120)
            ->setOffsetY(80);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getLegend()->getFont()->setItalic(true);
        */

        $objPHPPresentation->removeSlideByIndex(0);

// Set Style
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

// Generate sample data for chart
        echo date('H:i:s') . ' Generate sample data for chart' . EOL;
        $seriesData  = [
            'A' => 12,
            'B' => 15,
            'C' => 13,
            'D' => 17,
            'E' => 14,
            'F' => 9,
            'G' => 7,
            'H' => 8,
            'I' => 8,
            'J' => 15,
            'K' => 16,
            'L' => 14,
            'M' => 19,
            'N' => 13,
            'O' => 20,
        ];
        $seriesData1 = [
            'A' => 11,
            'B' => 22,
            'C' => 33,
        ];

        // Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

        // Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart = new Line();
        $series    = new Series('Downloads', $seriesData);
        $series->setShowSeriesName(true);
        $series->setShowValue(true);
        $lineChart->addSeries($series);

        // Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart)' . EOL;
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit(3);
        $shape->getPlotArea()->getAxisY()->setMajorUnit(5);

        // Create templated slide

        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

        // Create a line chart (that should be inserted in a shape)
        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_YELLOW));
        $oOutline->setWidth(2);

        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart1 = clone $lineChart;
        $series1    = $lineChart1->getSeries();
        $series1[0]->setOutline($oOutline);
        $series1[0]->getMarker()->setSymbol(Marker::SYMBOL_DIAMOND);
        $series1[0]->getMarker()->setSize(7);
        $lineChart1->setSeries($series1);
        echo date('H:i:s') . ' Create a shape (chart1)' . EOL;
        echo date('H:i:s') . ' Differences with previous : Values on right axis and Legend hidden' . EOL;
        $shape1 = clone $shape;
        $shape1->getLegend()->setVisible(false);
        $shape1->setName('PHPPresentation Weekly Downloads');
        $shape1->getTitle()->setText('PHPPresentation Weekly Downloads');
        $shape1->getPlotArea()->setType($lineChart1);
        $shape1->getPlotArea()->getAxisY()->setFormatCode('#,##0');
        $currentSlide->addShape($shape1);


        /*
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        *
        */
        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart2 = clone $lineChart;
        $series2    = $lineChart2->getSeries();
        $series2[0]->getFont()->setSize(25);
        $series2[0]->getMarker()->setSymbol(Marker::SYMBOL_TRIANGLE);
        $series2[0]->getMarker()->setSize(10);
        $lineChart2->setSeries($series2);

// Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart2)' . EOL;
        echo date('H:i:s') . ' Differences with previous : Values on right axis and Legend hidden' . EOL;
        $shape2 = clone $shape;
        $shape2->getLegend()->setVisible(false);
        $shape2->setName('PHPPresentation Weekly Downloads');
        $shape2->getTitle()->setText('PHPPresentation Weekly Downloads');
        $shape2->getPlotArea()->setType($lineChart2);
        $shape2->getPlotArea()->getAxisY()->setFormatCode('#,##0');
        $currentSlide->addShape($shape2);

// Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide #3' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart3 = clone $lineChart;

        $oGridLines1 = new Gridlines();
        $oGridLines1->getOutline()->setWidth(10);
        $oGridLines1->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));

        $oGridLines2 = new Gridlines();
        $oGridLines2->getOutline()->setWidth(1);
        $oGridLines2->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));

// Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart3)' . EOL;
        echo date('H:i:s') . ' Feature : Gridlines' . EOL;
        $shape3 = clone $shape;
        $shape3->setName('Shape 3');
        $shape3->getTitle()->setText('Chart with Gridlines');
        $shape3->getPlotArea()->setType($lineChart3);
        $shape3->getPlotArea()->getAxisX()->setMajorGridlines($oGridLines1);
        $shape3->getPlotArea()->getAxisY()->setMinorGridlines($oGridLines2);
        $currentSlide->addShape($shape3);

// Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide #4' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart4 = clone $lineChart;

        $oOutlineAxisX = new Outline();
        $oOutlineAxisX->setWidth(2);
        $oOutlineAxisX->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutlineAxisX->getFill()->getStartColor()->setRGB('012345');

        $oOutlineAxisY = new Outline();
        $oOutlineAxisY->setWidth(5);
        $oOutlineAxisY->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutlineAxisY->getFill()->getStartColor()->setRGB('ABCDEF');

// Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart4)' . EOL;
        echo date('H:i:s') . ' Feature : Axis Outline' . EOL;
        $shape4 = clone $shape;
        $shape4->setName('Shape 4');
        $shape4->getTitle()->setText('Chart with Outline on Axis');
        $shape4->getPlotArea()->setType($lineChart4);
        $shape4->getPlotArea()->getAxisX()->setOutline($oOutlineAxisX);
        $shape4->getPlotArea()->getAxisX()->setTitleRotation(45);
        $shape4->getPlotArea()->getAxisY()->setOutline($oOutlineAxisY);
        $shape4->getPlotArea()->getAxisY()->setTitleRotation(135);
        $currentSlide->addShape($shape4);

//        $currentSlide = $objPHPPresentation->createSlide();
//        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
//        $lineChart4 = new Line();
//
//        $oOutlineAxisX = new Outline();
//        $oOutlineAxisX->setWidth(2);
//        $oOutlineAxisX->getFill()->setFillType(Fill::FILL_SOLID);
//        $oOutlineAxisX->getFill()->getStartColor()->setRGB('012345');
//
//        $oOutlineAxisY = new Outline();
//        $oOutlineAxisY->setWidth(5);
//        $oOutlineAxisY->getFill()->setFillType(Fill::FILL_SOLID);
//        $oOutlineAxisY->getFill()->getStartColor()->setRGB('ABCDEF');
//
//// Create a shape (chart)
//        echo date('H:i:s') . ' Create a shape (chart4)' . EOL;
//        echo date('H:i:s') . ' Feature : Axis Outline' . EOL;
//        $shape4 = $currentSlide->createChartShape();
//        $shape4->setName('Shape 4');
//        $shape4->getTitle()->setText('Chart with Outline on Axis');
//        $shape4->getPlotArea()->setType($lineChart4);
//        $shape4->getPlotArea()->getAxisX()->setOutline($oOutlineAxisX);
//        $shape4->getPlotArea()->getAxisX()->setTitleRotation(45);
//        $shape4->getPlotArea()->getAxisY()->setOutline($oOutlineAxisY);
//        $shape4->getPlotArea()->getAxisY()->setTitleRotation(135);
//        $currentSlide->addShape($shape4);

        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }


    public function xy() {
        $objPHPPresentation = new PhpPresentation();
        $objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 07 Title')
            ->setSubject('Sample 07 Subject')
            ->setDescription('Sample 07 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // 设置背景宽高
        $objPHPPresentation->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_CUSTOM, true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(720, DocumentLayout::UNIT_PIXEL);

        $objPHPPresentation->removeSlideByIndex(0);
        $slide = $objPHPPresentation->createSlide();

        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));

        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

        $oOutline = new Outline();
        // 散点图线条类型
        $oOutline->getFill()->setFillType(Fill::FILL_NONE);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_DARKRED));
        $oOutline->setWidth(2);

        // Generate sample data for chart
        echo date('H:i:s') . ' Generate sample data for chart' . PHP_EOL;
        $seriesData  = [
            'A' => 12,
            'B' => 15,
            'C' => 13,
            'D' => 17,
            'E' => 14,
            'F' => 9,
            'G' => 7,
            'H' => 8,
            'I' => 8,
            'J' => 15,
            'K' => 16,
        ];
        $seriesData1 = [
            'A' => 15,
            'B' => 16,
            'C' => 32,
            'D' => 24,
            'E' => 16,
            'F' => 12,
            'G' => 24,
            'H' => 15,
            'I' => 14,
            'J' => 18,
            'K' => 22,
        ];
        // Create templated slide
        echo PHP_EOL . date('H:i:s') . ' Create templated slide' . PHP_EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);
        // Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . PHP_EOL;
        $lineChart = new Line();
        $series    = new Series('', $seriesData);
        $series->setShowSeriesName(false);
        $series->setShowValue(true);
        $lineChart->addSeries($series);
        $series2 = new Series('', $seriesData1);
        $series2->setShowSeriesName(false);
        $series2->setShowValue(true);
        $lineChart->addSeries($series2);

        // 散点图线条类型
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_DARKRED));
        $oOutline->setWidth(2);
        $series = $lineChart->getSeries();
        $series[0]->setOutline($oOutline);
        $series[0]->getMarker()->setSymbol(Marker::SYMBOL_DIAMOND);
        $series[0]->getMarker()->setSize(7);
        $font = $series[0]->getFont();
        $color = new Color('FFFFFFFF'); // 假设背景色为白色
        $font->getColor()->setRGB($color->getRGB());
        $lineChart->setSeries($series);

//        // 柱状图
//        $barChart = new Bar();
//        $barChart->setGapWidthPercent(158);
//        $series1 = new Series('2009', $seriesData);
//        $series1->setShowSeriesName(true);
//        $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4F81BD'));
//        $series1->getFont()->getColor()->setRGB('00FF00');
//        $series1->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
//        $series2 = new Series('2010', $seriesData1);
//        $series2->setShowSeriesName(true);
//        $series2->getFont()->getColor()->setRGB('FF0000');
//        $series2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFC0504D'));
//        $series2->setLabelPosition(Series::LABEL_INSIDEEND);
//        $barChart->addSeries($series1);
//        $barChart->addSeries($series2);

        // Create a shape (chart)
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)
            ->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
//        $shape->setShadow($oShadow);
//        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit(1);
        $shape->getPlotArea()->getAxisY()->setMajorUnit(5);
        // 给X轴加上边线
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        // 给Y轴加上边线
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_YELLOW));

//        $shape1 = clone $shape;
//        $shape1->getPlotArea()->setType($barChart);
//        $shape1->getView3D()->setRotationX(30);
//        $shape1->getView3D()->setPerspective(30);
//        $shape1->getLegend()->getBorder()->setLineStyle(Border::DASH_SOLID);
//        $shape1->getLegend()->getFont()->setItalic(true);
//        $shape1->getPlotArea()->getAxisX()->setMajorUnit(1);
//        $shape1->getPlotArea()->getAxisY()->setMajorUnit(5);
//        // 给X轴加上边线
//        $shape1->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
//        // 给Y轴加上边线
//        $shape1->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_YELLOW));


        // Create a line chart (that should be inserted in a shape)
        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    // XY散点图
    public function xyOK() {
        $objPHPPresentation = new PhpPresentation();
        $objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 07 Title')
            ->setSubject('Sample 07 Subject')
            ->setDescription('Sample 07 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // 设置背景宽高
        $objPHPPresentation->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_CUSTOM, true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(720, DocumentLayout::UNIT_PIXEL);

        $objPHPPresentation->removeSlideByIndex(0);
        $slide = $objPHPPresentation->createSlide();

        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFFFFFFF'));

        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

        $oOutline = new Outline();
        // 散点图线条类型
        $oOutline->getFill()->setFillType(Fill::FILL_NONE);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_DARKRED));
        $oOutline->setWidth(2);

        // Generate sample data for chart
        echo date('H:i:s') . ' Generate sample data for chart' . PHP_EOL;
        $seriesData  = [
            'A' => 12,
            'B' => 15,
            'C' => 13,
            'D' => 17,
            'E' => 14,
            'F' => 9,
            'G' => 7,
            'H' => 8,
            'I' => 8,
            'J' => 15,
            'K' => 16,
        ];
        $seriesData1 = [
            'A' => 11,
            'B' => 22,
            'C' => 33,
        ];
        // Create templated slide
        echo PHP_EOL . date('H:i:s') . ' Create templated slide' . PHP_EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);
        // Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . PHP_EOL;
        $lineChart = new Line();
        $series    = new Series('Downloads', $seriesData);
        $series->setShowSeriesName(false);
        $series->setShowValue(true);
        $lineChart->addSeries($series);
        $series2 = new Series('s2', $seriesData1);
        $series2->setShowSeriesName(false);
        $series2->setShowValue(true);
        $lineChart->addSeries($series2);

        // 散点图线条类型
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_DARKRED));
        $oOutline->setWidth(2);
        $series = $lineChart->getSeries();
        $series[0]->setOutline($oOutline);
        $series[0]->getMarker()->setSymbol(Marker::SYMBOL_DIAMOND);
        $series[0]->getMarker()->setSize(7);
        $lineChart->setSeries($series);

        // Create a shape (chart)
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)
            ->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit(1);
        $shape->getPlotArea()->getAxisY()->setMajorUnit(5);
        // 给X轴加上边线
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        // 给Y轴加上边线
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_YELLOW));


        // Create a line chart (that should be inserted in a shape)
        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function index2323() {
        $objPHPPresentation = new PhpPresentation();
        $objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 07 Title')
            ->setSubject('Sample 07 Subject')
            ->setDescription('Sample 07 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // 设置背景宽高
        $objPHPPresentation->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_CUSTOM, true)
            ->setCX(1280, DocumentLayout::UNIT_PIXEL)
            ->setCY(720, DocumentLayout::UNIT_PIXEL);

        $objPHPPresentation->removeSlideByIndex(0);
        $slide      = $objPHPPresentation->createSlide();
        $colorBlack = new Color('55555555');
        $slide->setBackground();

        $currentGroup = $objPHPPresentation->getActiveSlide()->createGroup();
        // Create a shape (text)
        $shape = $currentGroup->createRichTextShape()
            ->setHeight(200)
            ->setWidth(1180)
            ->setOffsetX(0)
            ->setOffsetY(200);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape->createTextRun('2');
        $textRun->getFont()->setBold(true)
            ->setSize(115)
            ->setName('黑体')
            ->setColor(new Color('FF40405C'));

        $shape1 = $currentGroup->createRichTextShape();
        $shape1->setHeight(30)
            ->setWidth(1180)
            ->setOffsetX(0)
            ->setOffsetY(450);
        $shape1->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $textRun = $shape1->createTextRun('微博平台传播链路分析');
        $textRun->getFont()->setBold(true)
            ->setSize(20)
            ->setName('黑体')
            ->setColor(new Color('FF40405C'));

        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    function index12() {
        echo date('H:i:s') . ' Create new PHPPresentation object' . EOL;
        $objPHPPresentation = new PhpPresentation();

// Set properties
        echo date('H:i:s') . ' Set properties' . EOL;
        $objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 07 Title')
            ->setSubject('Sample 07 Subject')
            ->setDescription('Sample 07 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

// Remove first slide
        echo date('H:i:s') . ' Remove first slide' . EOL;
        $objPHPPresentation->removeSlideByIndex(0);

// Set Style
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

        // Generate sample data for chart
        echo date('H:i:s') . ' Generate sample data for chart' . EOL;
        $seriesData = [
            'Monday'    => 12,
            'Tuesday'   => 15,
            'Wednesday' => 13,
            'Thursday'  => 17,
            'Friday'    => 14,
            'Saturday'  => 9,
            'Sunday'    => 7,
        ];

        // Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

        // Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a area chart (that should be inserted in a chart shape)' . EOL;
        $areaChart = new Area();
        $series    = new Series('Downloads', $seriesData);
        $series->setShowSeriesName(true);
        $series->setShowValue(true);
        $series->getFill()->setStartColor(new Color('ffffffff'));
        $series->setLabelPosition(Series::LABEL_INSIDEEND);
        $areaChart->addSeries($series);

        // Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart)' . EOL;
        $shape = $currentSlide->createChartShape();
        $shape->getTitle()->setVisible(false);
        $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(300)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($areaChart);
        $shape->getPlotArea()->getAxisX()->setTitle('Axis X');
        // 给X轴加上边线
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        // 给Y轴加上边线
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_YELLOW));

        $shape->getPlotArea()->getAxisY()->setTitle('Axis Y');
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getLegend()->getFont()->setItalic(true);
        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }


    function createTemplatedSlide(PhpPresentation $objPHPPresentation) {
        // Create slide
        $slide = $objPHPPresentation->createSlide();

        // Add logo
        $shape = $slide->createDrawingShape();
        $shape->setName('PHPPresentation logo')
            ->setDescription('PHPPresentation logo')
            ->setPath(WEB_PATH . '/uploads/static/images/bear.jpg')
            ->setHeight(36)
            ->setOffsetX(10)
            ->setOffsetY(10);
        $shape->getShadow()->setVisible(true)
            ->setDirection(45)
            ->setDistance(10);

        // Return slide
        return $slide;
    }


    public function index1(Request $request) { // 2.创建ppt对象
        $objPHPPowerPoint = new PhpPresentation();

        // 3.设置属性
        $objPHPPowerPoint->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 02 Title')
            ->setSubject('Sample 02 Subject')
            ->setDescription('Sample 02 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // 4.删除第一页(多页最好删除)
        $objPHPPowerPoint->removeSlideByIndex(0);


        //根据需求 调整for循环
        for ($i = 1; $i <= 3; $i++) {
            if ($i <= 2) {
                //创建幻灯片并添加到这个演示中
                $slide = $objPHPPowerPoint->createSlide();

                //创建一个形状(图)
                $shape = $slide->createDrawingShape();
                $shape->setName('内容图片name')
                    ->setDescription('内容图片 描述')
                    ->setPath(WEB_PATH . 'uploads/static/images/background.jpeg')
                    ->setResizeProportional(false)
                    ->setHeight(720)
                    ->setWidth(960);

                //创建一个形状(文本)
                $shape = $slide->createRichTextShape()
                    ->setHeight(60)
                    ->setWidth(960)
                    ->setOffsetX(10)
                    ->setOffsetY(50);
                $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
                $textRun = $shape->createTextRun('以后这个就是标题了');
                $textRun->getFont()->setBold(true)
                    ->setSize(20)
                    ->setColor(new Color('FFE06B20'));


                // 创建一个形状(文本)
                $shape = $slide->createRichTextShape()
                    ->setHeight(60)
                    ->setWidth(960)
                    ->setOffsetX()
                    ->setOffsetY(700);
                $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
                $textRun = $shape->createTextRun('时间:2017年10月19号');
                $textRun->getFont()->setBold(true)
                    ->setSize(10)
                    ->setColor(new Color('FFE06B20'));
            } else {
                // 创建一个幻灯片
                $slide = $objPHPPowerPoint->getActiveSlide();
                // 创建一个XY散点图
            }
        }

        $oWriterPPTX = IOFactory::createWriter($objPHPPowerPoint, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
        exit;
    }

    public function t() {
        $objPHPPresentation = new PhpPresentation();

// Set properties
        echo date('H:i:s') . ' Set properties' . EOL;
        $objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')->setLastModifiedBy('PHPPresentation Team')->setTitle('Sample 07 Title')->setSubject('Sample 07 Subject')->setDescription('Sample 07 Description')->setKeywords('office 2007 openxml libreoffice odt php')->setCategory('Sample Category');

// Remove first slide
        echo date('H:i:s') . ' Remove first slide' . EOL;
        $objPHPPresentation->removeSlideByIndex(0);

// Set Style
        $oFill = new Fill();
        $oFill->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));

        $oShadow = new Shadow();
        $oShadow->setVisible(true)->setDirection(45)->setDistance(10);

// Generate sample data for chart
        echo date('H:i:s') . ' Generate sample data for chart' . EOL;
        $seriesData = [
            'Monday 01' => 12,
            'Tuesday 02' => 15,
            'Wednesday 03' => 13,
            'Thursday 04' => 17,
            'Friday 05' => 14,
            'Saturday 06' => 9,
            'Sunday 07' => 7,
            'Monday 08' => 8,
            'Tuesday 09' => 8,
            'Wednesday 10' => 15,
            'Thursday 11' => 16,
            'Friday 12' => 14,
            'Saturday 13' => 14,
            'Sunday 14' => 13,
        ];

// Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart = new Line();
        $series = new Series('Downloads', $seriesData);
        $series->setShowSeriesName(true);
        $series->setShowValue(true);
        $lineChart->addSeries($series);

// Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart)' . EOL;
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Daily Downloads')->setResizeProportional(false)->setHeight(550)->setWidth(700)->setOffsetX(120)->setOffsetY(80);
        $shape->setShadow($oShadow);
        $shape->setFill($oFill);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText('PHPPresentation Daily Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit(3);
        $shape->getPlotArea()->getAxisY()->setMajorUnit(5);

// Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
        $oOutline = new Outline();
        $oOutline->getFill()->setFillType(Fill::FILL_SOLID);
        $oOutline->getFill()->setStartColor(new Color(Color::COLOR_YELLOW));
        $oOutline->setWidth(2);

        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart1 = clone $lineChart;
        $series1 = $lineChart1->getSeries();
        $series1[0]->setOutline($oOutline);
        $series1[0]->getMarker()->setSymbol(Marker::SYMBOL_DIAMOND);
        $series1[0]->getMarker()->setSize(7);
        $lineChart1->setSeries($series1);

// Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart1)' . EOL;
        echo date('H:i:s') . ' Differences with previous : Values on right axis and Legend hidden' . EOL;
        $shape1 = clone $shape;
        $shape1->getLegend()->setVisible(false);
        $shape1->setName('PHPPresentation Weekly Downloads');
        $shape1->getTitle()->setText('PHPPresentation Weekly Downloads');
        $shape1->getPlotArea()->setType($lineChart1);
        $shape1->getPlotArea()->getAxisY()->setFormatCode('#,##0');
        $currentSlide->addShape($shape1);

// Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

// Create a line chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a line chart (that should be inserted in a chart shape)' . EOL;
        $lineChart2 = clone $lineChart;
        $series2 = $lineChart2->getSeries();
        $series2[0]->getFont()->setSize(25);
        $series2[0]->getMarker()->setSymbol(Marker::SYMBOL_TRIANGLE);
        $series2[0]->getMarker()->setSize(10);
        $lineChart2->setSeries($series2);

// Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart2)' . EOL;
        echo date('H:i:s') . ' Differences with previous : Values on right axis and Legend hidden' . EOL;
        $shape2 = clone $shape;
        $shape2->getLegend()->setVisible(false);
        $shape2->setName('PHPPresentation Weekly Downloads');
        $shape2->getTitle()->setText('PHPPresentation Weekly Downloads');
        $shape2->getPlotArea()->setType($lineChart2);
        $shape2->getPlotArea()->getAxisY()->setFormatCode('#,##0');
        $currentSlide->addShape($shape2);

        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
        exit;
    }

    public function txt() {
        $objPHPPresentation = new PhpPresentation();
        echo date('H:i:s') . ' Set properties' . EOL;
        $objPHPPresentation->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')
            ->setTitle('Sample 01 Title')
            ->setSubject('Sample 01 Subject')
            ->setDescription('Sample 01 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        $currentSlide = $objPHPPresentation->getActiveSlide();

        for ($inc = 1; $inc <= 4; ++$inc) {
            // Create a shape (text)
            echo date('H:i:s') . ' Create a shape (rich text)' . EOL;
            $shape = $currentSlide->createRichTextShape()
                ->setHeight(200)
                ->setWidth(300);
            if (1 == $inc || 3 == $inc) {
                $shape->setOffsetX(10);
            } else {
                $shape->setOffsetX(320);
            }
            if (1 == $inc || 2 == $inc) {
                $shape->setOffsetY(10);
            } else {
                $shape->setOffsetY(220);
            }
            $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

            switch ($inc) {
                case 1:
                    $shape->getFill()->setFillType(Fill::FILL_NONE);
                    break;
                case 2:
                    $shape->getFill()->setFillType(Fill::FILL_GRADIENT_LINEAR)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF000000'));
                    break;
                case 3:
                    $shape->getFill()->setFillType(Fill::FILL_GRADIENT_PATH)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF000000'));
                    break;
                case 4:
                    $shape->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color('FF4672A8'))->setEndColor(new Color('FF4672A8'));
                    break;
            }

            $textRun = $shape->createTextRun('Use PHPPresentation!');
            $textRun->getFont()->setBold(true)->setSize(30)->setColor(new Color('FFE06B20'));
        }
        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
        exit;
    }

    public function line() {
        $objPHPPresentation = new PhpPresentation();
        // Create templated slide
        echo EOL . date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $this->createTemplatedSlide($objPHPPresentation);

        // Generate sample data for first chart
        echo date('H:i:s') . ' Generate sample data for chart' . EOL;
        $series1Data = ['Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293];
        $series2Data = ['Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379];
        $series3Data = ['Jan' => 233, 'Feb' => 146, 'Mar' => 238, 'Apr' => 175, 'May' => 108, 'Jun' => 257, 'Jul' => 199, 'Aug' => 201, 'Sep' => 88, 'Oct' => 147, 'Nov' => 287, 'Dec' => 105];

        // Create a bar chart (that should be inserted in a shape)
        echo date('H:i:s') . ' Create a stacked bar chart (that should be inserted in a chart shape)' . EOL;
        $StackedBarChart = new Bar();
        $series1 = new Series('2009', $series1Data);
        $series1->setShowSeriesName(false);
        $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4F81BD'));
        $series1->getFont()->getColor()->setRGB('00FF00');
        $series1->setShowValue(true);
        $series1->setShowPercentage(false);
        $series2 = new Series('2010', $series2Data);
        $series2->setShowSeriesName(false);
        $series2->getFont()->getColor()->setRGB('FF0000');
        $series2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFC0504D'));
        $series2->setShowValue(true);
        $series2->setShowPercentage(false);
        $series3 = new Series('2011', $series3Data);
        $series3->setShowSeriesName(false);
        $series3->getFont()->getColor()->setRGB('FF0000');
        $series3->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF804DC0'));
        $series3->setShowValue(true);
        $series3->setShowPercentage(false);
        $StackedBarChart->addSeries($series1);
        $StackedBarChart->addSeries($series2);
        $StackedBarChart->addSeries($series3);
        $StackedBarChart->setBarGrouping(Bar::GROUPING_STACKED);
        // Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart)' . EOL;
        $shape = $currentSlide->createChartShape();
        $shape->setName('PHPPresentation Monthly Downloads')
            ->setResizeProportional(false)
            ->setHeight(550)
            ->setWidth(700)
            ->setOffsetX(120)
            ->setOffsetY(80);
        $shape->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getTitle()->setText('PHPPresentation Monthly Downloads');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $shape->getPlotArea()->getAxisX()->setTitle('Month');
        $shape->getPlotArea()->getAxisY()->setTitle('Downloads');
        $shape->getPlotArea()->setType($StackedBarChart);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_SINGLE);
        $shape->getLegend()->getFont()->setItalic(true);
        $StackedBarChart->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);

        $oWriterPPTX = IOFactory::createWriter($objPHPPresentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }
}
