<?php

namespace App\Http\Controllers;

use App\Services\Converter;
use App\Services\PowerPointService;
use Illuminate\Foundation\Auth\Access\AuthorizesRequests;
use Illuminate\Foundation\Bus\DispatchesJobs;
use Illuminate\Foundation\Validation\ValidatesRequests;
use Illuminate\Http\Request;
use PhpOffice\Common\Drawing;
use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
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
use PhpOffice\PhpPresentation\Style\Bullet;
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

class PowerPointController extends Controller {
    use AuthorizesRequests, DispatchesJobs, ValidatesRequests;

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

        // 柱状图
        $barChart = new Bar();
        $barChart->setGapWidthPercent(158);
        $series1 = new Series('2009', $seriesData);
        $series1->setShowSeriesName(true);
        $series1->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4F81BD'));
        $series1->getFont()->getColor()->setRGB('00FF00');
        $series1->getDataPointFill(2)->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFE06B20'));
        $series2 = new Series('2010', $seriesData1);
        $series2->setShowSeriesName(true);
        $series2->getFont()->getColor()->setRGB('FF0000');
        $series2->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFC0504D'));
        $series2->setLabelPosition(Series::LABEL_INSIDEEND);
        $barChart->addSeries($series1);
        $barChart->addSeries($series2);

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

        $shape1 = clone $shape;
        $shape1->getPlotArea()->setType($barChart);
        $shape1->getView3D()->setRotationX(30);
        $shape1->getView3D()->setPerspective(30);
        $shape1->getLegend()->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape1->getLegend()->getFont()->setItalic(true);
        $shape1->getPlotArea()->getAxisX()->setMajorUnit(1);
        $shape1->getPlotArea()->getAxisY()->setMajorUnit(5);
        // 给X轴加上边线
        $shape1->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_BLUE));
        // 给Y轴加上边线
        $shape1->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_YELLOW));


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

    public function six() {
        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);
        // 生成标题文本
        $text       = '阶段热搜话题类型变化：节日热搜短暂盛行，社会经济议题恒久';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);
        $text       = '每个阶段均有相关话题资源位且均有热搜话题达到最高排名1，双节高热期，双节相关话题活跃度最高，前后阶段热搜量适中；双节相关热搜上榜当天大概率登上最高热搜排名。名1，双节高热期，双节相关话题活跃度最高，前后阶段热搜量适中；双节相关热搜上榜当天大概率登上最高热搜排名。';
        $properties = [
            'width'     => 1210,
            'height'    => 30,
            'x'         => 40,
            'y'         => 35,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
        ];
        $powerPoint->richText($slide, $text, $properties);

        $powerPoint->richText($slide, $text, $properties);
        $text       = '微博热搜话题在榜时长TOP20';
        $properties = [
            'width'     => 300,
            'height'    => 30,
            'x'         => 550,
            'y'         => 95,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $table  = $powerPoint->table($slide, 8, 270, 1150, 50, 135);
        $row    = $powerPoint->tableRow($table, 50, Fill::FILL_SOLID, 'FF4684D3');
        $prop   = [
            'fontColor' => Color::COLOR_WHITE,
            'bold'      => true,
            'fontSize'  => 12,
        ];
        $header = ['序号', '上榜时间', '热搜', '主持人', '在榜时长（小时）', '最高热度', '话题类型', '最高排名'];
        $width  = [60, 140, 350, 125, 145, 120, 120, 120];
        foreach ($header as $h => $head) {
            $powerPoint->cell($row, $head, $prop);
            $table->getRow(0)->getCell($h)->setWidth($width[$h]);
        }
        $data       = [
            [1, '央视新闻', '央媒', '1.32亿', '1648.6', 17, '', ''],
            [2, '节日君', '新浪官方', '632.4万', '607.0', 11, '', ''],
            [3, '新华社', '央媒', '1.1亿', '308.7', 11, '', ''],
            [4, '西部决策', '媒体', '525.6万', '732.3', 11, '', ''],
            [5, '人民日报', '央媒', '1.53亿', '411.2', 8, '', ''],
            [6, '央视网', '央媒', '1893.8万', '243.3', 8, '', ''],
            [7, '中国新闻网', '央媒', '7908.1万', '330.1', 6, '', ''],
            [8, '闪电视频', '媒体', '105.1万', '138.7', 5, '', ''],
            [9, '新华网', '央媒', '9672.5万', '113.7', 4, '', ''],
            [10, '四川观察', '媒体', '1118万', '206.1', 4, '', ''],
            [11, '中国蓝新闻', '媒体', '442万', '304.6', 4, '', ''],
            [12, 'CCTV4', '央媒', '648万', '165.6', 4, '', ''],
            [13, '成都发布', '媒体', '1379.2万', '119.6', 3, '', ''],
            [14, '中国消防', '政务', '968.6万', '45.6', 3, '', ''],
            [15, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [16, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [17, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [18, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [19, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [20, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
        ];
        $prop       = [
            'bold'     => false,
            'fontSize' => 12,
        ];
        $fillType   = Fill::FILL_NONE;
        $backColor  = Color::COLOR_WHITE;
        $cellBorder = [];
        $length     = count($data);
        foreach ($data as $key => $value) {
            if ($key == $length - 1) {
                $cellBorder = [
                    'bottom'      => 'FF4684D3',
                    'bottomWidth' => 2
                ];
            }
            $tableRow = $powerPoint->tableRow($table, 25, $fillType, $backColor);
            foreach ($value as $k => $v) {
                $powerPoint->cell($tableRow, $v, $prop, $cellBorder);
            }
        }

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function seven() {
        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);
        // 生成标题文本
        $text       = '热搜热度TOP20：出游、调休等社经话题热度高';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);
        $text       = '每个阶段均有相关话题资源位且均有热搜话题达到最高排名1，双节高热期，双节相关话题活跃度最高，前后阶段热搜量适中；双节相关热搜上榜当天大概率登上最高热搜排名。名1，双节高热期，双节相关话题活跃度最高，前后阶段热搜量适中；双节相关热搜上榜当天大概率登上最高热搜排名。';
        $properties = [
            'width'     => 1210,
            'height'    => 30,
            'x'         => 40,
            'y'         => 35,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
        ];
        $powerPoint->richText($slide, $text, $properties);

        $powerPoint->richText($slide, $text, $properties);
        $text       = '微博热搜热度TOP20';
        $properties = [
            'width'     => 300,
            'height'    => 30,
            'x'         => 550,
            'y'         => 95,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $table  = $powerPoint->table($slide, 8, 270, 1150, 50, 135);
        $row    = $powerPoint->tableRow($table, 50, Fill::FILL_SOLID, 'FF4684D3');
        $prop   = [
            'fontColor' => Color::COLOR_WHITE,
            'bold'      => true,
            'fontSize'  => 12,
        ];
        $header = ['序号', '时间', '热搜', '主持人', '在榜时长（小时）', '最高热度', '话题类型', '最高排名'];
        $width  = [60, 140, 350, 125, 145, 120, 120, 120];
        foreach ($header as $h => $head) {
            $powerPoint->cell($row, $head, $prop);
            $table->getRow(0)->getCell($h)->setWidth($width[$h]);
        }
        $data       = [
            [1, '央视新闻', '央媒', '1.32亿', '1648.6', 17, '', ''],
            [2, '节日君', '新浪官方', '632.4万', '607.0', 11, '', ''],
            [3, '新华社', '央媒', '1.1亿', '308.7', 11, '', ''],
            [4, '西部决策', '媒体', '525.6万', '732.3', 11, '', ''],
            [5, '人民日报', '央媒', '1.53亿', '411.2', 8, '', ''],
            [6, '央视网', '央媒', '1893.8万', '243.3', 8, '', ''],
            [7, '中国新闻网', '央媒', '7908.1万', '330.1', 6, '', ''],
            [8, '闪电视频', '媒体', '105.1万', '138.7', 5, '', ''],
            [9, '新华网', '央媒', '9672.5万', '113.7', 4, '', ''],
            [10, '四川观察', '媒体', '1118万', '206.1', 4, '', ''],
            [11, '中国蓝新闻', '媒体', '442万', '304.6', 4, '', ''],
            [12, 'CCTV4', '央媒', '648万', '165.6', 4, '', ''],
            [13, '成都发布', '媒体', '1379.2万', '119.6', 3, '', ''],
            [14, '中国消防', '政务', '968.6万', '45.6', 3, '', ''],
            [15, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [16, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [17, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [18, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [19, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
            [20, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
        ];
        $prop       = [
            'bold'     => false,
            'fontSize' => 12,
        ];
        $fillType   = Fill::FILL_NONE;
        $backColor  = Color::COLOR_WHITE;
        $cellBorder = [];
        $length     = count($data);
        foreach ($data as $key => $value) {
            if ($key == $length - 1) {
                $cellBorder = [
                    'bottom'      => 'FF4684D3',
                    'bottomWidth' => 2
                ];
            }
            $tableRow = $powerPoint->tableRow($table, 25, $fillType, $backColor);
            foreach ($value as $k => $v) {
                $powerPoint->cell($tableRow, $v, $prop, $cellBorder);
            }
        }

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function eight() {
        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);
        // 生成标题文本
        $text       = '热搜主持贡献TOP账号：媒体主导双节热搜话题走向';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $shape = $slide->createRichTextShape();
        $shape->setHeight(600)->setWidth(1210)->setOffsetX(40)->setOffsetY(35);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->getActiveParagraph()->getFont()->setSize(14)
            ->setName('微软雅黑')
            ->setColor(new Color('FF000000'));
        $shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createTextRun('媒体尤其是央媒主导双节热搜走向，除微博官方节日君主动参与主持热搜话题外，所有TOP贡献均为媒体；');

        $shape->createParagraph()->getAlignment()->setLevel(0)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->createTextRun('热搜热度引导动作主要通过节日文化、明星效应、热点话题投票互动、参与活动抽奖等方式带动话题热度，如转发接力点亮国旗图标、汉服文化、AI再现古人对话等节日文化；明星视频/运动员祝福视频等。');
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createParagraph()->getAlignment()->setLevel(1)
            ->setMarginLeft(75)
            ->setIndent(-25);

        $text       = '微博热搜主持人贡献TOP10';
        $properties = [
            'width'     => 300,
            'height'    => 30,
            'x'         => 550,
            'y'         => 120,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $table  = $powerPoint->table($slide, 7, 1200, 270, 50, 158);
        $row    = $powerPoint->tableRow($table, 50, Fill::FILL_SOLID, 'FF4684D3');
        $prop   = [
            'fontColor' => Color::COLOR_WHITE,
            'bold'      => true,
            'fontSize'  => 12,
        ];
        $header = ['序号', '账号', '类型', '粉丝数', '贡献热度值（万）', '贡献热搜个数', '主要动作'];
        $width  = [50, 130, 110, 100, 130, 120, 560];
        foreach ($header as $h => $head) {
            $powerPoint->cell($row, $head, $prop);
            $table->getRow(0)->getCell($h)->setWidth($width[$h]);
        }
        $data       = [
            [1, '央视新闻', '央媒', '1.32亿', '1648.6', 17, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [2, '节日君', '新浪官方', '632.4万', '607.0', 1, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [3, '新华社', '央媒', '1.1亿', '308.7', 11, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [4, '西部决策', '媒体', '525.6万', '732.3', 11, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [5, '人民日报', '央媒', '1.53亿', '411.2', 8, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [6, '央视网', '央媒', '1893.8万', '243.3', 8, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [7, '中国新闻网', '央媒', '7908.1万', '330.1', 6, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [8, '闪电视频', '媒体', '105.1万', '138.7', 5, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [9, '新华网', '央媒', '9672.5万', '113.7', 4, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
            [10, '四川观察', '媒体', '1118万', '206.1', 4, '转发接力点亮国旗图标、国庆/中秋晚会明星演唱视频带播放、转发中秋晚会节目单等'],
        ];
        $prop       = [
            'bold'     => false,
            'fontSize' => 11,
        ];
        $cellBorder = [];
        $length     = count($data);
        $fillType   = Fill::FILL_NONE;
        $backColor  = Color::COLOR_WHITE;
        foreach ($data as $key => $value) {
            if ($key == $length - 1) {
                $cellBorder = [
                    'bottom'      => 'FF4684D3',
                    'bottomWidth' => 2
                ];
            }
            $tableRow = $powerPoint->tableRow($table, 40, $fillType, $backColor);
            foreach ($value as $k => $v) {
                $powerPoint->cell($tableRow, $v, $prop, $cellBorder);
            }
        }

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
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

    public function ppt() {
        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);
        // 生成标题文本
        $text       = '官方热搜资源机制—资源位热搜约占双节热搜5%，官方主持热搜推荐约26%';
        $properties = [
            'width'     => 1000,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $shape = $slide->createRichTextShape();
        $shape->setHeight(600)->setWidth(1210)->setOffsetX(40)->setOffsetY(35);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->getActiveParagraph()->getFont()->setSize(14)
            ->setName('微软雅黑')
            ->setColor(new Color('FF000000'));
        $shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createTextRun('监测期内，官方提供热搜资源位13个，其中置顶热搜2个，辟谣热搜1个，推荐双节热搜10个，资源位热搜在双节相关热搜中约占5%；');

        $shape->createParagraph()->getAlignment()->setLevel(0)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->createTextRun('微博官方各类账号主持热搜34个，在双节相关热搜中约占比13%，其中官方主持热搜推荐9个，推荐比例约为26%；');
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createParagraph()->getAlignment()->setLevel(0)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->createTextRun('预热阶段微博官方主持推荐2个资源位，双节期间推荐6个资源位，消散期推荐主持1个资源位，官方贡献比例逐步下降，在第一阶段和第二阶段重点引导，第二阶段贡献资源位比例最高。');
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createParagraph()->getAlignment()->setLevel(1)
            ->setMarginLeft(75)
            ->setIndent(-25);

        $x    = range(1, 10);
        $data = [];
        foreach ($x as $y) {
            for ($i = 0; $i < 10; $i++) {
                $data[$y][$i] = rand(10, 99);
            }
        }
        $data = [
            '2014-01-01' => [33956],
            '2014-01-02' => [76345],
            '2014-01-03' => [140000],
            '2014-01-07' => [22222],
            '2014-01-06' => [33956],
        ];
        foreach ($data as &$value) {
            foreach ($value as $key => $item) {
                $value[strtotime($key)] = $item;
                unset($value[$key]);
            }
        }
        $data       = [
            '新华社'  => [
                '2014-01-01' => 12,
                '2014-01-02' => NULL,
            ],
            '新华社A' => [
                '2014-01-02' => 44,
            ],
            '新华社B' => [
                '2014-01-03' => 33,
            ],
            '新华社C' => [
                '2014-01-04' => 22,
            ],
            '新华社D' => [
                '2014-01-05' => 66,
            ],
            '新华社E' => [
                '2014-01-06' => 37,
            ],
        ];
        $properties = [
            'showSeriesName' => true,
            'showValue'      => true
        ];
        $pointAttr  = [
            'borderColor' => ['FF4472C4', 'FFED7D31', 'FFA5A5A5'],
            'pointColor'  => ['FF4472C4', 'FFED7D31', 'FFA5A5A5'],
        ];
        // 图线属性
        $lineChart = $powerPoint->XYScatter($data, $properties, $pointAttr);

        $oGridLines1 = new Gridlines();
        $oGridLines1->getOutline()->setWidth(10);
        $oGridLines1->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));
        $shape = $slide->createChartShape();
        $shape->setName('')
            ->setResizeProportional(false)->setHeight(280)->setWidth(1180)->setOffsetX(40)->setOffsetY(130);
        $shape->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getTitle()->setText('');
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit();
        $shape->getPlotArea()->getAxisY()->setMajorUnit(2);
        $shape->getLegend()->setVisible(false);
        $shape->setName('');
        $shape->getTitle()->setText('');
        $shape->getPlotArea()->setType($lineChart);
        $shape->getPlotArea()->getAxisX()->setMajorTickMark(Axis::TICK_MARK_INSIDE);
        $shape->getPlotArea()->getAxisX()->setMajorGridlines($oGridLines1);
        $shape->getPlotArea()->getAxisX()->setTitle('');
        $shape->getPlotArea()->getAxisY()->setTitle('');
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));


        $sources   = [
            [
                'title' => '平台双节热搜',
                'data'  => [53, 177, 32],
            ],
            [
                'title' => '官方贡献热搜',
                'data'  => [10, 21, 3]
            ],
            [
                'title' => '总资源位热搜',
                'data'  => [2, 10, 1],
            ],
            [
                'title' => '官方资源位推荐',
                'data'  => [2, 6, 1]
            ]
        ];
        $data      = [];
        $colors    = ['FF4472C4', 'FFED7D31', 'FFA5A5A5'];
        $names     = ['第一阶段', '第二阶段', '第三阶段'];
        $tableData = [];
        foreach ($names as $key => $name) {
            foreach ($sources as $source) {
                $data[$key]['data'][] = $source['data'][$key] ?? 0;
                $data[$key]['name']   = $name;
                $data[$key]['attr']   = [
                    'type'  => Fill::FILL_SOLID,
                    'color' => $colors[$key],
                ];
                // 生成表格数据
                $reverse           = array_reverse($source['data']);
                $tableData[$key][] = $reverse[$key];
            }
        }
        $barChart = $powerPoint->barChart($data, [], true);
        $barChart->setBarGrouping(Bar::GROUPING_STACKED);
        // 第一个图表高度
        $shape = $slide->createChartShape();
        $shape->setName('')
            ->setResizeProportional(false)
            ->setHeight(280) // 1-140 2-170 3-200
            ->setWidth(450)
            ->setOffsetX(100)
            ->setOffsetY(325);
        $shape->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getTitle()->setText('');
        $shape->getTitle()->getFont()->setItalic(false);
        $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $shape->getPlotArea()->setType($barChart);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getLegend()->getFont()->setItalic(false);
//        $shape->getLegend()->setVisible(false);
        $shape->getLegend()->setPosition(Chart\Legend::POSITION_LEFT);
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisY()->setTitle('');
        $shape->getPlotArea()->getAxisY()->setIsVisible(false);
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisX()->setTitle('');
        $shape->getPlotArea()->getAxisX()->setMajorUnit(1);
        $shape->getPlotArea()->setOffsetX(100);
        $shape->getPlotArea()->setOffsetY(-10);
        $shape->getLegend()->getFont()->setColor(new Color('FFA5A5A5'));
        $shape->getPlotArea()->getAxisX()->setIsVisible(false);
        $shape->setRotation(50);
        $shape->getLegend()->setPosition(Chart\Legend::POSITION_LEFT);
        $barChart->setBarGrouping(Bar::GROUPING_STACKED);
        // 右下图表
        $barChart = $powerPoint->barChart($data, [], true, true);
        $barChart->setBarGrouping(Bar::GROUPING_STACKED);
        // 第一个图表高度
        $shape = $slide->createChartShape();
        $shape->setName('')
            ->setResizeProportional(false)
            ->setHeight(345) // 1-140 2-170 3-200
            ->setWidth(630)
            ->setOffsetX(590)
            ->setOffsetY(350);
        $shape->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getTitle()->setText('');
        $shape->getTitle()->getFont()->setItalic(false);
        $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $shape->getPlotArea()->setType($barChart);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getLegend()->getFont()->setItalic(false);
//        $shape->getLegend()->setVisible(false);
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisY()->setTitle('');
        $shape->getPlotArea()->getAxisY()->setIsVisible(false);
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisX()->setTitle('');
        $shape->getPlotArea()->getAxisX()->setMajorUnit(1);
        $shape->getPlotArea()->setOffsetX(100);
        $shape->getPlotArea()->setOffsetY(-10);
        $shape->getLegend()->getFont()->setColor(new Color('FFA5A5A5'));
        $shape->getPlotArea()->getAxisX()->setIsVisible(false);
        $shape->setRotation(50);
        $shape->getLegend()->setPosition(Chart\Legend::POSITION_TOP);
        $barChart->setBarGrouping(Bar::GROUPING_STACKED);
        // 生成左下表格
        $table  = $powerPoint->table($slide, 4, 470, 50, 155, 600);
        $row    = $powerPoint->tableRow($table, 20, Fill::FILL_NONE);
        $prop   = [
            'fontColor' => 'FF595959',
            'fontSize'  => 9,
        ];
        $header = array_column($sources, 'title');
        $width  = [105, 105, 105, 105];
        foreach ($header as $h => $head) {
            $powerPoint->cell($row, $head, $prop);
            $table->getRow(0)->getCell($h)->setWidth($width[$h]);
        }
        $prop       = [
            'bold'     => false,
            'fontSize' => 9,
        ];
        $fillType   = Fill::FILL_NONE;
        $backColor  = Color::COLOR_WHITE;
        $cellBorder = [];
        foreach ($tableData as $key => $vv) {
            $tableRow = $powerPoint->tableRow($table, 20, $fillType, $backColor);
            foreach ($vv as $k => $v) {
                $powerPoint->cell($tableRow, $v, $prop, $cellBorder);
            }
        }

        $properties = [
            'width'     => 900,
            'height'    => 30,
            'x'         => 50,
            'y'         => 685,
            'fontColor' => 'FF40405C',
            'fontName'  => '微软雅黑',
            'fontSize'  => 9,
        ];
        // 生成富文本
        $powerPoint->richText($slide, '注：本页官方贡献热搜指官方主持的热搜（共34个），资源位指微博置顶、推荐和辟谣的热搜，官方资源位推荐指的是推荐热搜资源位中官方支持的热搜。', $properties);

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function ten() {
        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);
        // 生成标题文本
        $text       = '最热热搜传播路径—超级月亮：媒体和KOL助推历史热搜翻红';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $shape = $slide->createRichTextShape();
        $shape->setHeight(600)->setWidth(1210)->setOffsetX(40)->setOffsetY(35);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->getActiveParagraph()->getFont()->setSize(14)
            ->setName('微软雅黑')
            ->setColor(new Color('FF000000'));
        $shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createTextRun('超级月亮热搜有三个类似主题热搜在当天和次日登上热搜，超级月亮上榜最早，在榜时长最长，一是话题较宽泛，主题发散性强，二是为历史热搜，曾早在2022年7月和2023年8月多次登上过热搜，在中秋这个特殊节点再次爆发热度登上热搜；');

        $shape->createParagraph()->getAlignment()->setLevel(0)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->createTextRun('再次助推超级月亮上热搜的主要账号为媒体和生活关联性博主，9个主要账号中央视新闻、人民日报等媒体账号4个，游戏博主“逆水寒手游”、美妆博主“澜晓溪vivi”和旅游摄影博主“大斌堡”KOL账号3个，此外明星账号和网评员账号各1个。');
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createParagraph()->getAlignment()->setLevel(1)
            ->setMarginLeft(75)
            ->setIndent(-25);

        $text       = '微博热搜主持人贡献TOP15';
        $properties = [
            'width'     => 300,
            'height'    => 30,
            'x'         => 550,
            'y'         => 130,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $table  = $powerPoint->table($slide, 6, 1180, 270, 50, 158);
        $row    = $powerPoint->tableRow($table, 50, Fill::FILL_SOLID, 'FF4684D3');
        $prop   = [
            'fontColor' => Color::COLOR_WHITE,
            'bold'      => true,
            'fontSize'  => 12,
        ];
        $header = ['上榜时间', '热搜', '主持人', '最高排名', '最高热度（万）', '在榜时长（小时）'];
        $width  = [210, 340, 155, 125, 175, 175];
        foreach ($header as $h => $head) {
            $powerPoint->cell($row, $head, $prop);
            $table->getRow(0)->getCell($h)->setWidth($width[$h]);
        }
        $data       = [
            [1, '央视新闻', '央媒', '1.32亿', '1648.6', 17],
            [2, '节日君', '新浪官方', '632.4万', '607.0', 1],
            [3, '新华社', '央媒', '1.1亿', '308.7', 11],
//            [4, '西部决策', '媒体', '525.6万', '732.3', 11, '', ''],
//            [5, '人民日报', '央媒', '1.53亿', '411.2', 8, '', ''],
//            [6, '央视网', '央媒', '1893.8万', '243.3', 8, '', ''],
//            [7, '中国新闻网', '央媒', '7908.1万', '330.1', 6, '', ''],
//            [8, '闪电视频', '媒体', '105.1万', '138.7', 5, '', ''],
//            [9, '新华网', '央媒', '9672.5万', '113.7', 4, '', ''],
//            [10, '四川观察', '媒体', '1118万', '206.1', 4, '', ''],
//            [11, '中国蓝新闻', '媒体', '442万', '304.6', 4, '', ''],
//            [12, 'CCTV4', '央媒', '648万', '165.6', 4, '', ''],
//            [13, '成都发布', '媒体', '1379.2万', '119.6', 3, '', ''],
//            [14, '中国消防', '政务', '968.6万', '45.6', 3, '', ''],
//            [15, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
        ];
        $prop       = [
            'bold'     => false,
            'fontSize' => 12,
        ];
        $cellBorder = [];
        $length     = count($data);
        foreach ($data as $key => $value) {
            if ($key == $length - 1) {
                $cellBorder = [
                    'bottom'      => 'FF4684D3',
                    'bottomWidth' => 2
                ];
            }
            if ($key == 0) {
                $fillType  = Fill::FILL_SOLID;
                $backColor = 'FFFBE5D6';
            } else {
                $fillType  = Fill::FILL_NONE;
                $backColor = Color::COLOR_WHITE;
            }
            $tableRow = $powerPoint->tableRow($table, 40, $fillType, $backColor);
            foreach ($value as $k => $v) {
                $powerPoint->cell($tableRow, $v, $prop, $cellBorder);
            }
        }
        $x    = range(1, 10);
        $data = [];
        foreach ($x as $y) {
            for ($i = 0; $i < 10; $i++) {
                $data[$y][$i] = rand(10, 99);
            }
        }
        $data = [
            '2014-01-01' => [33956],
            '2014-01-02' => [76345],
            '2014-01-03' => [140000],
            '2014-01-07' => [22222],
            '2014-01-06' => [33956],
        ];
        foreach ($data as &$value) {
            foreach ($value as $key => $item) {
                $value[strtotime($key)] = $item;
                unset($value[$key]);
            }
        }
        $data       = [
            '新华社' => [
                '2014-01-01' => 12,
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
                '',
            ],
        ];
        $properties = [
            'showSeriesName' => true,
            'showValue'      => true
        ];
        $pointAttr  = [
            'borderColor' => ['FF4472C4', 'FFED7D31', 'FFA5A5A5'],
            'pointColor'  => ['FF4472C4', 'FFED7D31', 'FFA5A5A5'],
        ];
        // 图线属性
        $lineChart = $powerPoint->XYScatter($data, $properties, $pointAttr);

        $oGridLines1 = new Gridlines();
        $oGridLines1->getOutline()->setWidth(10);
        $oGridLines1->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));

        $oGridLines2 = new Gridlines();
        $oGridLines2->getOutline()->setWidth(1);
        $oGridLines2->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color(Color::COLOR_DARKGREEN));

        $shape = $slide->createChartShape();
        $shape->setName('')
            ->setResizeProportional(false)->setHeight(370)->setWidth(1180)->setOffsetX(40)->setOffsetY(350);
        $shape->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getTitle()->setText('');
        $shape->getPlotArea()->setType($lineChart);
        $shape->getView3D()->setRotationX(30);
        $shape->getView3D()->setPerspective(30);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit();
        $shape->getPlotArea()->getAxisY()->setMajorUnit(2);
        $shape->getLegend()->setVisible(false);
        $shape->setName('');
        $shape->getTitle()->setText('');
        $shape->getPlotArea()->setType($lineChart);
        $shape->getPlotArea()->getAxisX()->setMajorTickMark(Axis::TICK_MARK_INSIDE);
        $shape->getPlotArea()->getAxisX()->setMajorGridlines($oGridLines1);
//        $shape->getPlotArea()->getAxisY()->setMinorGridlines($oGridLines2);
        $shape->getPlotArea()->getAxisX()->setTitle('');
        $shape->getPlotArea()->getAxisY()->setTitle('');
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));

        $properties     = [
            'width'     => 600,
            'height'    => 30,
            'x'         => 50,
            'y'         => 670,
            'fontColor' => 'FF40405C',
            'fontName'  => '微软雅黑',
            'fontSize'  => 9,
        ];
        $textProperties = [
            'fontColor' => 'FFFF0000',
            'fontName'  => '微软雅黑',
            'fontSize'  => 9,
        ];
        // 生成富文本
        $richText = $powerPoint->richText($slide, '注：本页仅展示互动量', $properties);
        $richText = $powerPoint->richText($slide, '超1万', $textProperties, $richText);
        $powerPoint->richText($slide, ' 的账号信息（共9篇）。', $properties, $richText);

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function nine() {
        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);
        // 生成标题文本
        $text       = '阶段热搜话题类型变化：节日热搜短暂盛行，社会经济议题恒久';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);
        $text       = '微博热搜主持人贡献TOP15';
        $properties = [
            'width'     => 300,
            'height'    => 30,
            'x'         => 550,
            'y'         => 125,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        $shape = $slide->createRichTextShape();
        $shape->setHeight(600)->setWidth(1210)->setOffsetX(40)->setOffsetY(35);
        $shape->getActiveParagraph()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_LEFT)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->getActiveParagraph()->getFont()->setSize(14)
            ->setName('微软雅黑')
            ->setColor(new Color('FF000000'));
        $shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET);
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createTextRun('讨论。节日相关的话题在其特定时期内热度较高，但迅速消散，中秋话题在预热期的热度最高，国庆话，节日话题最后都回归社会和经济发展讨论。');

        $shape->createParagraph()->getAlignment()->setLevel(0)
            ->setMarginLeft(25)
            ->setIndent(-25);
        $shape->createTextRun('节日相关的话题在其特定时期内热度较高，但迅速消散，中秋话题在预热期的热度最高，国庆话，节日话题最后都回归社会和经济发展讨论。');
        $shape->getActiveParagraph()->getBulletStyle()->setBulletChar('•');
        $shape->createParagraph()->getAlignment()->setLevel(1)
            ->setMarginLeft(75)
            ->setIndent(-25);

        $table  = $powerPoint->table($slide, 8, 270, 1180, 50, 155);
        $row    = $powerPoint->tableRow($table, 50, Fill::FILL_SOLID, 'FF4684D3');
        $prop   = [
            'fontColor' => Color::COLOR_WHITE,
            'bold'      => true,
            'fontSize'  => 12,
        ];
        $header = ['序号', '账号', '类型', '粉丝（万）', '贡献热度值（万）', '贡献热搜个数', '主持热搜阶段划分', '话题类型划分'];
        $width  = [50, 90, 90, 110, 135, 112, 300, 330];
        foreach ($header as $h => $head) {
            $powerPoint->cell($row, $head, $prop);
            $table->getRow(0)->getCell($h)->setWidth($width[$h]);
        }
        $data       = [
            [1, '央视新闻', '央媒', '1.32亿', '1648.6', 17, '', ''],
            [2, '节日君', '新浪官方', '632.4万', '607.0', 11, '', ''],
            [3, '新华社', '央媒', '1.1亿', '308.7', 11, '', ''],
            [4, '西部决策', '媒体', '525.6万', '732.3', 11, '', ''],
            [5, '人民日报', '央媒', '1.53亿', '411.2', 8, '', ''],
            [6, '央视网', '央媒', '1893.8万', '243.3', 8, '', ''],
            [7, '中国新闻网', '央媒', '7908.1万', '330.1', 6, '', ''],
            [8, '闪电视频', '媒体', '105.1万', '138.7', 5, '', ''],
            [9, '新华网', '央媒', '9672.5万', '113.7', 4, '', ''],
            [10, '四川观察', '媒体', '1118万', '206.1', 4, '', ''],
            [11, '中国蓝新闻', '媒体', '442万', '304.6', 4, '', ''],
            [12, 'CCTV4', '央媒', '648万', '165.6', 4, '', ''],
            [13, '成都发布', '媒体', '1379.2万', '119.6', 3, '', ''],
            [14, '中国消防', '政务', '968.6万', '45.6', 3, '', ''],
            [15, '农民频道', '媒体', '344.6万', '267.3', 3, '', ''],
        ];
        $prop       = [
            'bold'     => false,
            'fontSize' => 12,
        ];
        $fillType   = Fill::FILL_NONE;
        $backColor  = Color::COLOR_WHITE;
        $cellBorder = [];
        $length     = count($data);
        foreach ($data as $key => $value) {
            if ($key == $length - 1) {
                $cellBorder = [
                    'bottom'      => 'FF4684D3',
                    'bottomWidth' => 2
                ];
            }
            $tableRow = $powerPoint->tableRow($table, 30, $fillType, $backColor);
            foreach ($value as $k => $v) {
                $powerPoint->cell($tableRow, $v, $prop, $cellBorder);
            }
        }


//        $series1Data = ['Jan' => 133, 'Feb' => 99, 'Mar' => 191, 'Apr' => 205, 'May' => 167, 'Jun' => 201, 'Jul' => 240, 'Aug' => 226, 'Sep' => 255, 'Oct' => 264, 'Nov' => 283, 'Dec' => 293];
//        $series2Data = ['Jan' => 266, 'Feb' => 198, 'Mar' => 271, 'Apr' => 305, 'May' => 267, 'Jun' => 301, 'Jul' => 340, 'Aug' => 326, 'Sep' => 344, 'Oct' => 364, 'Nov' => 383, 'Dec' => 379];
//        $series3Data = ['Jan' => 233, 'Feb' => 146, 'Mar' => 238, 'Apr' => 175, 'May' => 108, 'Jun' => 257, 'Jul' => 199, 'Aug' => 201, 'Sep' => 88, 'Oct' => 147, 'Nov' => 287, 'Dec' => 105];
        $number = count($data);
        $loop   = 15;
        $data   = [];
        $colors = ['FF4472C4', 'FFFFC000', 'FFA5A5A5'];
        $names  = ['第一阶段', '第二阶段', '第三阶段'];
        $length = $number;
        for ($i = 0; $i < 3; $i++) {
            for ($j = 0; $j < $length; $j++) {
                $p                             = $j + 0;
                $data[$i]['data'][$p]          = rand(10, 999);
                $data[$i]['name']              = $names[$i];
                $data[$i]['attr']['type']      = Fill::FILL_SOLID;
                $data[$i]['attr']['color']     = $colors[$i];
                $data[$i]['attr']['fontColor'] = 'FF828282';
            }
        }
        $barChart = $powerPoint->barChart($data, [], true);
        $barChart->setBarGrouping(Bar::GROUPING_STACKED);
        // 第一个图表高度
        $height = 140 + 30 * ($number - 1);
        $shape  = $slide->createChartShape();
        $shape->setName('')
            ->setResizeProportional(false)
            ->setHeight($height) // 1-140 2-170 3-200
            ->setWidth(320)
            ->setOffsetX(650)
            ->setOffsetY(196);
        $shape->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getTitle()->setText('');
        $shape->getTitle()->getFont()->setItalic(false);
        $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $shape->getPlotArea()->setType($barChart);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getLegend()->getFont()->setItalic(false);
//        $shape->getLegend()->setVisible(false);
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisY()->setTitle('');
        $shape->getPlotArea()->getAxisY()->setIsVisible(false);
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisX()->setTitle('');
        $shape->getPlotArea()->getAxisX()->setMajorUnit(1);
        $shape->getLegend()->getFont()->setColor(new Color('FFA5A5A5'));
        $shape->getPlotArea()->getAxisX()->setIsVisible(false);
        $shape->getLegend()->setPosition(Chart\Legend::POSITION_BOTTOM);

        $barChart->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
        $barChart->setBarGrouping(Bar::GROUPING_STACKED);


        $data   = [];
        $colors = ['FF4472C4', 'FFFFC000', 'FFA5A5A5', 'FFD442BF', 'FF0ABA5F'];
        $names  = ['国庆', '中秋', '社会', '清明', '春节'];
        $length = $number;
        for ($i = 0; $i < 3; $i++) {
            for ($j = 0; $j < $length; $j++) {
                $data[$i]['name']          = $names[$i];
                $data[$i]['data'][$j]      = rand(10, 999);
                $data[$i]['attr']['type']  = Fill::FILL_SOLID;
                $data[$i]['attr']['color'] = $colors[$i];
            }
        }
        $barChart = $powerPoint->barChart($data, [], true);
        $shape    = $slide->createChartShape();
        $height   = 107 + 30 * ($number - 1);
        $shape->setName('')
            ->setResizeProportional(false)
            ->setHeight($height)  // 1-107
            ->setWidth(320)
            ->setOffsetX(950)
            ->setOffsetY(196);
        $shape->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getTitle()->setText('');
        $shape->getTitle()->getFont()->setItalic(false);
        $shape->getTitle()->getAlignment()->setHorizontal(Alignment::HORIZONTAL_RIGHT);
        $shape->getPlotArea()->setType($barChart);
        $shape->getLegend()->getBorder()->setLineStyle(Border::LINE_NONE);
        $shape->getLegend()->getFont()->setItalic(false);
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisY()->setTitle('');
        $shape->getPlotArea()->getAxisY()->setIsVisible(false);
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_NONE);
        $shape->getPlotArea()->getAxisX()->setTitle('');
        $shape->getPlotArea()->getAxisX()->setMajorUnit(1.2);
        $shape->getPlotArea()->getAxisX()->setIsVisible(false);
        $shape->getLegend()->getFont()->setColor(new Color('FFA5A5A5'));
        $shape->getLegend()->setOffsetX(800);
        $shape->getLegend()->setOffsetY(600);
        $barChart->setBarDirection(Bar3D::DIRECTION_HORIZONTAL);
        $barChart->setBarGrouping(Bar::GROUPING_STACKED);
        // Create a shape (chart)
        echo date('H:i:s') . ' Create a shape (chart)' . EOL;


        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function five() {
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
        $seriesData2 = [
            'A' => 22,
            'B' => 33,
            'C' => 55,
            'D' => 34,
            'E' => 44,
            'F' => 22,
            'G' => 35,
            'H' => 21,
            'I' => 64,
            'J' => 24,
            'K' => 43,
        ];
        $seriesData3 = [
            'A' => 147,
            'B' => 245,
            'C' => 210,
            'D' => 159,
            'E' => 173,
            'F' => 166,
            'G' => 300,
            'H' => 114,
            'I' => 139,
            'J' => 182,
            'K' => 263,
        ];
        // 数据和属性
        $data = [
            [
                'data' => $seriesData,
                'name' => '话题消散期',
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FFFFC000',
                    ''
                ]
            ],
            [
                'data' => $seriesData1,
                'name' => '话题预热期',
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FF70AD47',
                ]
            ],
            [
                'data' => $seriesData2,
                'name' => '话题话题期',
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FF5B9BD5'
                ]
            ],
            [
                'data' => $seriesData3,
                'name' => '总计',
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FF43682B'
                ]
            ]
        ];

        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);
        $table = $powerPoint->table($slide, 4, 270, 1180, 50, 100);
        // Add row
//        $row = $shape->createRow();
//        $row->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FF4684D3'))->setEndColor(new Color('FF4684D3'));
//        $row->setHeight(45);
        $row    = $powerPoint->tableRow($table, 45, Fill::FILL_SOLID, 'FF4684D3');
        $prop   = [
            'fontColor' => Color::COLOR_WHITE,
            'bold'      => true,
            'fontSize'  => 14,
        ];
        $header = ['类型', '话题预热期', '双节话题期', '话题消散期'];
        foreach ($header as $head) {
            $powerPoint->cell($row, $head, $prop);
        }
        $data       = [
            ['国庆', '18.87%', '23.73%', '3.13%'],
            ['中秋', '39.62%', '32.20%', '0.00%'],
            ['中秋国庆双节', '9.43%', '4.52%', '0.00%'],
            ['经济', '1.89%', '2.82%', '9.38%'],
            ['旅游', '5.66%', '7.34%', '18.75%'],
            ['明星网红', '3.77%', '5.65%', '3.13%'],
            ['社会', '20.75%', '23.73%', '65.63%'],
            ['阶段热搜个数', '53', '177', '32'],
        ];
        $count      = count($data);
        $prop       = [
            'bold'     => false,
            'fontSize' => 12,
        ];
        $fillType   = Fill::FILL_NONE;
        $backColor  = Color::COLOR_WHITE;
        $cellBorder = [];
        foreach ($data as $key => $value) {
            if ($key == ($count - 1)) {
                $fillType     = Fill::FILL_SOLID;
                $backColor    = 'FFEEEEEE';
                $cellBorder   = [
                    'top'      => 'FF4684D3',
                    'topWidth' => 2
                ];
                $prop['bold'] = true;
            }
            $tableRow = $powerPoint->tableRow($table, 27, $fillType, $backColor);
            foreach ($value as $k => $v) {
                $powerPoint->cell($tableRow, $v, $prop, $cellBorder);
            }
        }

        $powerPoint->image($slide, WEB_PATH . 'uploads/static/images/image1.png', 410, 400, 80, 330);
        $powerPoint->image($slide, WEB_PATH . 'uploads/static/images/image2.png', 410, 400, 460, 330);
        $powerPoint->image($slide, WEB_PATH . 'uploads/static/images/image3.png', 410, 400, 830, 330);


        echo date('H:i:s') . ' Create templated slide' . EOL;
        $currentSlide = $presentation->createSlide();

        // Create a shape (text)
//        echo date('H:i:s') . ' Create a shape (rich text) with list with red bullet' . EOL;
//        $shape = $currentSlide->createRichTextShape();
//        $shape->setHeight(100);
//        $shape->setWidth(400);
//        $shape->setOffsetX(100);
//        $shape->setOffsetY(100);
//        $shape->getActiveParagraph()->getBulletStyle()->setBulletType(Bullet::TYPE_BULLET)->setBulletColor(new Color(Color::COLOR_RED));
//
//        $shape->createText('Alpha');
//        $shape->createParagraph()->createText('Beta');
//        $shape->createParagraph()->createText('Delta');
//        $shape->createParagraph()->createText('Epsilon');

//        $shape = $slide->createDrawingShape();
//        $shape->setName('')
//            ->setDescription('')
//            ->setPath( WEB_PATH . 'uploads/static/images/image1.png')
//            ->setResizeProportional(false)
//            ->setOffsetX(80)
//            ->setOffsetY(330)
//            ->setHeight(400)
//            ->setWidth(410);

//        $shape = $slide->createDrawingShape();
//        $shape->setName('')
//            ->setDescription('')
//            ->setPath( WEB_PATH . 'uploads/static/images/image2.png')
//            ->setResizeProportional(false)
//            ->setOffsetX(460)
//            ->setOffsetY(330)
//            ->setHeight(400)
//            ->setWidth(410);
//
//        $shape = $slide->createDrawingShape();
//        $shape->setName('')
//            ->setDescription('')
//            ->setPath( WEB_PATH . 'uploads/static/images/image3.png')
//            ->setResizeProportional(false)
//            ->setOffsetX(830)
//            ->setOffsetY(330)
//            ->setHeight(400)
//            ->setWidth(410);

        // 生成标题文本
        $text       = '阶段热搜话题类型变化：节日热搜短暂盛行，社会经济议题恒久';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);
        // 生成文本
        $text       = '节日相关的话题在其特定时期内热度较高，但迅速消散，中秋话题在预热期的热度最高，国庆话题在预热期和双节话题期的热度较高，但在话题消散期时，热度显著下降；而旅游、经济和社会热搜个数占比呈现上升趋势，节日话题最后都回归社会和经济发展讨论。';
        $properties = [
            'width'     => 1210,
            'height'    => 30,
            'x'         => 40,
            'y'         => 35,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
        ];
        $powerPoint->richText($slide, $text, $properties);

        // 生成文本
        $text       = '话题预热期热搜词条词云';
        $properties = [
            'width'     => 330,
            'height'    => 30,
            'x'         => 60,
            'y'         => 680,
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'alignment' => Alignment::HORIZONTAL_CENTER,
            'bold'      => true
        ];
        $powerPoint->richText($slide, $text, $properties);
        $text       = '双节话题期热搜词条词云';
        $properties = [
            'width'     => 330,
            'height'    => 30,
            'x'         => 480,
            'y'         => 680,
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'alignment' => Alignment::HORIZONTAL_CENTER,
            'bold'      => true
        ];
        $powerPoint->richText($slide, $text, $properties);
        $text       = '话题消散期热搜词条词云';
        $properties = [
            'width'     => 330,
            'height'    => 30,
            'x'         => 860,
            'y'         => 680,
            'fontName'  => '黑体',
            'fontSize'  => 12,
            'alignment' => Alignment::HORIZONTAL_CENTER,
            'bold'      => true
        ];
        $powerPoint->richText($slide, $text, $properties);

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function three() {
        $x    = range(1, 20);
        $data = [];
        foreach ($x as $y) {
            for ($i = 0; $i < 3; $i++) {
                $data[$y][$i] = rand(100, 999);
            }
        }

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
        $seriesData2 = [
            'A' => 25,
            'B' => 36,
            'C' => 12,
            'D' => 34,
            'E' => 36,
            'F' => 22,
            'G' => 44,
            'H' => 15,
            'I' => 44,
            'J' => 28,
            'K' => 32,
        ];
        $seriesData2 = [
            'A' => 22,
            'B' => 33,
            'C' => 55,
            'D' => 34,
            'E' => 44,
            'F' => 22,
            'G' => 35,
            'H' => 21,
            'I' => 64,
            'J' => 24,
            'K' => 43,
        ];
        $seriesData3 = [
            'A' => 147,
            'B' => 245,
            'C' => 210,
            'D' => 159,
            'E' => 173,
            'F' => 166,
            'G' => 300,
            'H' => 114,
            'I' => 139,
            'J' => 182,
            'K' => 263,
        ];
        // 数据和属性
        $data = [
            [
                'data' => $seriesData,
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FFFFC000',
                    ''
                ]
            ],
            [
                'data' => $seriesData1,
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FF70AD47',
                ]
            ],
            [
                'data' => $seriesData2,
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FF5B9BD5'
                ]
            ],
            [
                'data' => $seriesData3,
                'attr' => [
                    'line'      => Fill::FILL_SOLID,
                    'lineColor' => 'FF43682B'
                ]
            ]
        ];

        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;
        // 图线属性
        $lineChart = $powerPoint->XYLineChart($data);
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);

        $shape = $slide->createChartShape();
        $shape->setName()->setResizeProportional(false)
            ->setHeight(610)->setWidth(1250)->setOffsetX(2)->setOffsetY(92);
        // 设置隐藏系列标签
        $shape->getLegend()->setVisible(false);
//        $shape->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getTitle()->setText('');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit(1);
        $shape->getPlotArea()->getAxisY()->setMajorUnit(20);
        // X轴刻度线
//        $shape->getPlotArea()->getAxisX()->setMajorTickMark(Axis::TICK_MARK_INSIDE);
        // 给X轴加上边线
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));
        // 隐藏X轴名称
        $shape->getPlotArea()->getAxisX()->setTitle('');
        // 改变X轴字体颜色
        $shape->getPlotArea()->getAxisX()->getFont()->getColor()->setARGB(Color::COLOR_RED);
        // 给Y轴加上边线
//        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffd9d9d9'));
        // 隐藏Y轴名称
        $shape->getPlotArea()->getAxisY()->setTitle('');


        // 生成标题文本
        $text       = '热搜排名及最高排名趋势：符合假期前后工休闲暇规律';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);
        // 生成文本
        $text       = '每个阶段均有相关话题资源位且均有热搜话题达到最高排名1，双节高热期，双节相关话题活跃度最高，前后阶段热搜量适中；双节相关热搜上榜当天大概率登上最高热搜排名。';
        $properties = [
            'width'     => 1210,
            'height'    => 30,
            'x'         => 40,
            'y'         => 35,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
        ];
        $powerPoint->richText($slide, $text, $properties);

        // 生成文本
        $text       = '━';
        $properties = [
            'width'     => 1000,
            'height'    => 30,
            'x'         => 50,
            'y'         => 110,
            'fontColor' => 'FF70AD47',
            'fontName'  => '黑体',
            'bold'      => true,
            'fontSize'  => 18,
        ];
        // 生成文本
        $line           = '━';
        $lineProperties = [
            'fontColor' => 'FF5B9BD5',
            'fontName'  => '黑体',
            'bold'      => true,
            'fontSize'  => 18,
        ];
        $textProperties = [
            'fontColor' => 'FF000000',
            'fontName'  => '黑体',
            'fontSize'  => 13,
        ];
        // 生成富文本
        $richText = $powerPoint->richText($slide, $line, $properties);
        $richText = $powerPoint->richText($slide, ' 话题预热期 ', $textProperties, $richText);
        $richText = $powerPoint->richText($slide, $line, $lineProperties, $richText);
        $richText = $powerPoint->richText($slide, ' 话题话题期 ', $textProperties, $richText);
        // 设置话题消散期颜色
        $lineProperties['fontColor'] = 'FFFFC000';
        // 生成文本
        $richText = $powerPoint->richText($slide, $line, $lineProperties, $richText);
        $richText = $powerPoint->richText($slide, ' 话题消散期 ', $textProperties, $richText);
        // 设置总计颜色
        $lineProperties['fontColor'] = 'FF43682B';
        // 生成文本
        $richText = $powerPoint->richText($slide, $line, $lineProperties, $richText);
        $powerPoint->richText($slide, ' 总计', $textProperties, $richText);

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function four() {
        $x    = range(1, 20);
        $data = [];
        foreach ($x as $y) {
            for ($i = 0; $i < 20; $i++) {
                $data[$y][$i] = rand(100, 999);
            }
        }
//        dd($data);die();
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
        $powerPoint  = new PowerPointService();
//        $data         = [$seriesData, $seriesData1];
        $presentation = $powerPoint->phpPresentation;
        // 图线属性
        $lineChart = $powerPoint->XYScatter($data, [], []);
        // 创建幻灯片
        $slide = $powerPoint->slide($presentation);

        $shape = $slide->createChartShape();
        $shape->setName()->setResizeProportional(false)
            ->setHeight(370)->setWidth(1210)->setOffsetX(40)->setOffsetY(92);
        // 设置隐藏系列标签
        $shape->getLegend()->setVisible(false);
        // 设置标题名称
        $shape->getTitle()->setText('');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart);
        $shape->getLegend()->getFont()->setItalic(true);
        // X轴的步进
        $shape->getPlotArea()->getAxisX()->setMajorUnit(2);
        // Y轴的步进
        $shape->getPlotArea()->getAxisY()->setMajorUnit(100);
        // X轴刻度线
//        $shape->getPlotArea()->getAxisX()->setMajorTickMark(Axis::TICK_MARK_INSIDE);
        // 给X轴加上边线
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));
        // 隐藏X轴名称
        $shape->getPlotArea()->getAxisX()->setTitle('');
        // 改变X轴字体颜色
        $shape->getPlotArea()->getAxisX()->getFont()->getColor()->setARGB(Color::COLOR_RED);
        // 给Y轴加上边线
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('ffd9d9d9'));
        // 隐藏Y轴名称
        $shape->getPlotArea()->getAxisY()->setTitle('');
        // Y轴倒置
        $shape->getPlotArea()->getAxisY()->setIsReversedOrder(true);


        // 第二个图表
        $lineChart2 = $powerPoint->XYScatter($data, [], [], true);
        $shape      = $slide->createChartShape();
        $shape->setName()->setResizeProportional(false)
            ->setHeight(370)->setWidth(1210)->setOffsetX(40)->setOffsetY(400);
        // 设置隐藏系列标签
        $shape->getLegend()->setVisible(false);
        $shape->getTitle()->setText('');
        $shape->getTitle()->getFont()->setItalic(true);
        $shape->getPlotArea()->setType($lineChart2);
//        $shape->getLegend()->getBorder()->setLineStyle(Border::DASH_SOLID);
        $shape->getLegend()->getFont()->setItalic(true);
        $shape->getPlotArea()->getAxisX()->setMajorUnit(2);
        $shape->getPlotArea()->getAxisY()->setMajorUnit(100);
        // 给X轴加上边线
        $shape->getPlotArea()->getAxisX()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));
        // 隐藏X轴名称
        $shape->getPlotArea()->getAxisX()->setTitle('');
        // 改变X轴字体颜色
        $shape->getPlotArea()->getAxisX()->getFont()->setColor(new Color('FFFF0000'));        // 给Y轴加上边线
        $shape->getPlotArea()->getAxisY()->getOutline()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color('FFD9D9D9'));
        // 隐藏Y轴名称
        $shape->getPlotArea()->getAxisY()->setTitle('');
        // Y轴倒置
        $shape->getPlotArea()->getAxisY()->setIsReversedOrder(true);

        // 生成文本
        $text       = '热搜上榜时间和排名分布情况';
        $properties = [
            'width'     => 30,
            'height'    => 300,
            'x'         => 2,
            'y'         => 103,
            'backColor' => 'FF2A63F3',
            'fontName'  => '黑体',
            'fontSize'  => 14,
            'fontColor' => 'FFFFFFFF',
            'alignment' => Alignment::HORIZONTAL_CENTER
        ];
        $powerPoint->richText($slide, $text, $properties);

        // 生成文本
        $text       = '热搜最高热度和排名分布情况';
        $properties = [
            'width'     => 30,
            'height'    => 300,
            'x'         => 2,
            'y'         => 415,
            'backColor' => 'FF2A63F3',
            'fontName'  => '黑体',
            'fontSize'  => 14,
            'fontColor' => 'FFFFFFFF',
            'alignment' => Alignment::HORIZONTAL_CENTER
        ];
        $powerPoint->richText($slide, $text, $properties);

        // 生成标题文本
        $text       = '热搜排名及最高排名趋势：符合假期前后工休闲暇规律';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true
        ];
        $powerPoint->richText($slide, $text, $properties);
        // 生成文本
        $text       = '每个阶段均有相关话题资源位且均有热搜话题达到最高排名1，双节高热期，双节相关话题活跃度最高，前后阶段热搜量适中；双节相关热搜上榜当天大概率登上最高热搜排名。';
        $properties = [
            'width'     => 1210,
            'height'    => 30,
            'x'         => 40,
            'y'         => 35,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
        ];
        $powerPoint->richText($slide, $text, $properties);

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
        //路径 /uploads/ppt/  必须存在
        $path = WEB_PATH . 'uploads/ppt/';
        if (!file_exists($path)) {
            mkdir($path, 0777, true);
        }
        $file = $path . DIRECTORY_SEPARATOR . time() . '.pptx';
        $oWriterPPTX->save($file);
        var_dump($file);
    }

    public function two() {
        $powerPoint   = new PowerPointService();
        $presentation = $powerPoint->phpPresentation;

        $slide = $powerPoint->slide($presentation);
        $text       = '日热搜个数与热度走势：符合节日话题的时间性关注模式';
        $properties = [
            'width'     => 800,
            'height'    => 30,
            'x'         => 40,
            'y'         => 2,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'bold'      => true,
        ];
        $powerPoint->richText($slide, $text, $properties);

        // 生成文本
        $text       = '微博热搜分布';
        $properties = [
            'width'     => 40,
            'height'    => 500,
            'x'         => 2,
            'y'         => 180,
            'backColor' => 'FF2A63F3',
            'fontName'  => '黑体',
            'fontSize'  => 18,
            'fontColor' => 'FFFFFFFF',
            'alignment' => Alignment::HORIZONTAL_CENTER
        ];
        $richText = $powerPoint->richText($slide, $text, $properties);
        $richText->setInsetTop(170);

        // 生成文本
        $text       = '日热搜个数和对应热搜热度的变化趋势高度一致，均呈现明显的“低潮-高潮-回落”的节日讨论规律，与中秋国庆节日时间特点高度匹配体现了网民对节日话题的时间性关注模式。';
        $properties = [
            'width'     => 1210,
            'height'    => 30,
            'x'         => 40,
            'y'         => 35,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
        ];
        $richText = $powerPoint->richText($slide, $text, $properties);
        $richText->createBreak();
        $text       = '调休话题在双节前后热度较高，前期资源位引导消费主题，高热期资源位聚焦打卡和节日仪式感。';
        $properties = [
            'width'     => 1210,
            'height'    => 30,
            'x'         => 40,
            'y'         => 88,
            'backColor' => 'FFFFFFFF',
            'fontName'  => '黑体',
        ];
        $richText = $powerPoint->richText($slide, $text, $properties);
        $powerPoint->image($slide, WEB_PATH . 'uploads/static/images/image4.png', 1200, 580, 50, 120);

        $oWriterPPTX = IOFactory::createWriter($presentation, 'PowerPoint2007');
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
