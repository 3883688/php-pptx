<?php

namespace App\Services;

use PhpOffice\PhpPresentation\DocumentLayout;
use PhpOffice\PhpPresentation\Exception\FileNotFoundException;
use PhpOffice\PhpPresentation\Exception\OutOfBoundsException;
use PhpOffice\PhpPresentation\PhpPresentation;
use PhpOffice\PhpPresentation\Shape\Chart\Marker;
use PhpOffice\PhpPresentation\Shape\Chart\Series;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Bar;
use PhpOffice\PhpPresentation\Shape\Chart\Type\Line;
use \PhpOffice\PhpPresentation\Slide;
use PhpOffice\PhpPresentation\Style\Alignment;
use PhpOffice\PhpPresentation\Style\Border;
use PhpOffice\PhpPresentation\Style\Color;
use PhpOffice\PhpPresentation\Style\Fill;
use PhpOffice\PhpPresentation\Style\Outline;
use \PhpOffice\PhpPresentation\Shape\RichText\Run;
use  \PhpOffice\PhpPresentation\Shape\RichText;
use \PhpOffice\PhpPresentation\Shape\Table;
use \PhpOffice\PhpPresentation\Shape\Table\Row;
use \PhpOffice\PhpPresentation\Shape\Table\Cell;
use \PhpOffice\PhpPresentation\Shape\Drawing\File;

class PowerPointService {
    // 幻灯片宽度
    const WIDTH = 33.87;
    // 幻灯片高度
    const HEIGHT = 19.05;

    public $phpPresentation = null;

    public function __construct() {
        $this->phpPresentation = new PhpPresentation();
        $this->properties();
    }

    /**
     * XY散点图
     * @param         $data
     * @param array   $properties 散点图图表属性
     * @param array   $pointAttr  散点图点属性
     * @return Line
     */
    public function XYScatter($data, array $properties = [], array $pointAttr = []): Line {
        $lineChart      = new Line();
        $showSeriesName = $properties['showSeriesName'] ?? false;
        $showValue      = $properties['showValue'] ?? false;
        foreach ($data as $key => $value) {
            $series     = new Series($key, $value);
            $series->setShowSeriesName($showSeriesName);
            $series->setShowValue($showValue);
            $lineChart->addSeries($series);
        }
        // 折线类型
        $line = $properties['line'] ?? Fill::FILL_NONE;
        // 点颜色
        $color = $properties['color'] ?? 'FFFFC6A0';
        // 线宽度【像素】
        $width = $properties['width'] ?? 2;
        // 点类型
        $symbol = $properties['symbol'] ?? Marker::SYMBOL_CIRCLE;
        // 点大小【像素】
        $size = $properties['size'] ?? 6;

        // 散点图线条类型
        $outline      = $this->outline($line, $color, $width);
        $series       = $lineChart->getSeries();
        $borderColors = $pointAttr['borderColor'] ?? [];
        $pointColor   = $pointAttr['pointColor'] ?? [];
        foreach ($series as $key => &$serie) {
            $serie->setOutline($outline);
            // XY散点图属性配置
            // 设置点的边框颜色
            $borderColor = $borderColors[$key] ?? 'FFFF9045';
            $serie->getMarker()->getBorder()->setColor(new Color($borderColor));
            // 设置点的填充颜色
            $pointFillColor = $pointColor[$key] ?? $color;
            $serie->getMarker()->getFill()->setFillType(Fill::FILL_SOLID)->setStartColor(new Color($pointFillColor));
            $serie->getMarker()->setSymbol($symbol);
            $serie->getMarker()->setSize($size);
        }
        $lineChart->setSeries($series);
        return $lineChart;
    }

    /**
     * XY散点图|折线图
     * @param         $data
     * @param array   $properties
     * @return Line
     */
    public function XYLineChart($data, array $properties = []): Line {
        $lineChart      = new Line();
        $showSeriesName = $properties['showSeriesName'] ?? false;
        $showValue      = $properties['showValue'] ?? false;
        foreach ($data as $key => $value) {
            $series = new Series($key, $value['data']);
            $series->setShowSeriesName($showSeriesName);
            $series->setShowValue($showValue);
            $lineChart->addSeries($series);
        }
        // 折线类型
        $line = $properties['line'] ?? Fill::FILL_NONE;
        // 线宽度【像素】
        $width = $properties['width'] ?? 20;

        $series = $lineChart->getSeries();
        // 定义一个颜色数组
        foreach ($series as $k => &$serie) {
            $lineColor = $data[$k]['attr']['lineColor'];
            $outline   = $this->outline($line, $lineColor, $width);
            $serie->setOutline($outline);
            $serie->getOutline()->getFill()->setFillType(Fill::FILL_SOLID);
        }
        $lineChart->setSeries($series);
        return $lineChart;
    }

    /**
     * 创建PPT基本属性
     */
    public function properties(): PowerPointService {
        $this->phpPresentation->getDocumentProperties()->setCreator('PHPOffice')
            ->setLastModifiedBy('PHPPresentation Team')->setTitle('Sample 07 Title')
            ->setSubject('Sample 07 Subject')->setDescription('Sample 07 Description')
            ->setKeywords('office 2007 openxml libreoffice odt php')
            ->setCategory('Sample Category');

        // 设置PPT宽高
        $this->phpPresentation->getLayout()->setDocumentLayout(DocumentLayout::LAYOUT_A4, true)
            ->setCX(Converter::cmToPixel(self::WIDTH), DocumentLayout::UNIT_PIXEL)
            ->setCY(Converter::cmToPixel(self::HEIGHT), DocumentLayout::UNIT_PIXEL);
        return $this;
    }

    /**
     * 创建幻灯片
     */
    public function slide(PhpPresentation $phpPresentation): Slide {
        return $phpPresentation->createSlide();
    }

    /**
     * 边框线
     * @param string $line  边线类型
     * @param string $color 颜色类型
     * @param mixed  $width 宽度
     * @return Outline
     */
    public function outline(string $line = Fill::FILL_NONE, string $color = Color::COLOR_DARKRED, $width = 2): Outline {
        $oOutline = new Outline();
        // 散点图线条类型
        $oOutline->getFill()->setFillType($line);
        $oOutline->getFill()->setStartColor(new Color($color));
        $oOutline->setWidth($width);
        return $oOutline;
    }

    /**
     * 创建富文本
     * @param Slide         $slide
     * @param               $text       mixed 文字
     * @param               $properties array 文本属性
     * @param RichText|null $richText
     * @return RichText
     */
    public function richText(Slide $slide, $text, array $properties, RichText $richText = null): RichText {
        // 背景色【ARGB】
        $backColor = $properties['backColor'] ?? 'FFFFFFFF';
        // 字体名称
        $fontName = $properties['fontName'] ?? '';
        // 是否加粗
        $isBold = $properties['bold'] ?? false;
        // 字体大小
        $fontSize = $properties['fontSize'] ?? 14;
        // 字体颜色
        $fontColor = $properties['fontColor'] ?? 'FF000000';
        // 对齐方式
        $alignment = $properties['alignment'] ?? Alignment::HORIZONTAL_LEFT;
        // 创建富文本
        if (empty($richText)) {
            $richText = $slide->createRichTextShape()->setHeight($properties['height'])->setWidth($properties['width']);
        }
        if (isset($properties['x'])) {
            $richText->setOffsetX($properties['x']);
        }
        if (isset($properties['y'])) {
            $richText->setOffsetY($properties['y']);
        }
        $richText->getActiveParagraph()->getAlignment()->setHorizontal($alignment);
        $richText->getFill()->setFillType(Fill::FILL_SOLID)->setRotation(90)->setStartColor(new Color($backColor))->setEndColor(new Color($backColor));
        $textRun = $richText->createTextRun($text);
        $textRun->getFont()->setName($fontName)->setBold($isBold)->setSize($fontSize)->setColor(new Color($fontColor));
        $textRun->setLanguage('zh-cn');
        return $richText;
    }

    /**
     * 创建表格对象
     * @param Slide $slide
     * @param       $columns integer 列数
     * @param       $width   integer 宽度
     * @param       $height  integer 高度
     * @param       $x       integer X偏移量
     * @param       $y       integer Y偏移量
     * @return Table
     */
    public function table(Slide $slide, int $columns, int $width, int $height, int $x, int $y): Table {
        $shape = $slide->createTableShape($columns);
        $shape->setHeight($width);
        $shape->setWidth($height);
        $shape->setOffsetX($x);
        $shape->setOffsetY($y);
        return $shape;
    }

    /**
     * 生成表格行对象
     * @param Table  $table
     * @param int    $height    行高
     * @param string $fillType  填充类型
     * @param string $backColor 填充颜色
     * @return Row
     */
    public function tableRow(Table $table, int $height = 0, string $fillType = Fill::FILL_SOLID, string $backColor = Color::COLOR_BLACK): Row {
        $row = $table->createRow();
        $row->getFill()->setFillType($fillType)->setStartColor(new Color($backColor))->setEndColor(new Color($backColor));
        $row->setHeight($height);
        return $row;
    }

    /**
     * 生成单元格和单元格属性设置
     * @param Row   $row
     * @param mixed $text
     * @param array $properties
     * @param array $border
     * @return Cell
     */
    public function cell(Row $row, $text, array $properties = [], array $border = []): Cell {
        // 字体名称
        $fontName = $properties['fontName'] ?? '微软雅黑';
        // 是否加粗
        $isBold = $properties['bold'] ?? false;
        // 字体大小
        $fontSize = $properties['fontSize'] ?? 14;
        // 字体颜色
        $fontColor = $properties['fontColor'] ?? Color::COLOR_BLACK;
        // 对齐方式
        $alignment    = $properties['alignment'] ?? Alignment::HORIZONTAL_CENTER;
        $topBorder    = $border['top'] ?? 'FFE4E4E4';
        $bottomBorder = $border['bottom'] ?? 'FFE4E4E4';
        $leftBorder   = $border['left'] ?? 'FFE4E4E4';
        $rightBorder  = $border['right'] ?? 'FFE4E4E4';
        // 边框粗细
        $topBorderWidth    = $border['topWidth'] ?? 1;
        $bottomBorderWidth = $border['bottomWidth'] ?? 1;
        $leftBorderWidth   = $border['leftWidth'] ?? 1;
        $rightBorderWidth  = $border['rightWidth'] ?? 1;
        $oCell             = $row->nextCell();
        $textRun           = $oCell->createTextRun($text);
        $textRun->getFont()->setName($fontName)->setColor(new Color($fontColor))->setBold($isBold)->setSize($fontSize);
        // 设置上边框颜色
        $oCell->getBorders()->getTop()->setLineWidth($topBorderWidth)->setLineStyle(Border::LINE_SINGLE)->setColor(new Color($topBorder));
        // 设置下边框颜色
        $oCell->getBorders()->getBottom()->setLineWidth($bottomBorderWidth)->setLineStyle(Border::LINE_SINGLE)->setColor(new Color($bottomBorder));
        // 设置左边框颜色
        $oCell->getBorders()->getLeft()->setLineWidth($leftBorderWidth)->setLineStyle(Border::LINE_SINGLE)->setColor(new Color($leftBorder));
        // 设置右边框颜色
        $oCell->getBorders()->getRight()->setLineWidth($rightBorderWidth)->setLineStyle(Border::LINE_SINGLE)->setColor(new Color($rightBorder));
        $oCell->getActiveParagraph()->getAlignment()->setVertical($alignment)->setHorizontal($alignment);
        return $oCell;
    }

    /**
     * 生成图片
     * @param Slide  $slide
     * @param string $image  图片地址
     * @param int    $width  宽度
     * @param int    $height 高度
     * @param int    $x      X偏移量
     * @param int    $y      Y便宜
     * @param string $name   名称
     * @param string $desc   描述
     * @return File
     * @throws FileNotFoundException
     */
    public function image(Slide $slide, string $image, int $width, int $height, int $x, int $y, $name = '', $desc = ''): File {
        $shape = $slide->createDrawingShape();
        $shape->setName($name)
            ->setDescription($desc)
            ->setPath($image)
            ->setResizeProportional(false)
            ->setOffsetX($x)
            ->setOffsetY($y)
            ->setHeight($height)
            ->setWidth($width);
        return $shape;
    }

    /**
     * 柱形图生成
     * @param $data           array 数据
     * @param $fontStyle      array 柱形图字体样式
     * @param $showSeriesName bool 是否显示系列名称
     * @param $showValue      bool 是否显示数值
     * @param $showPercentage bool 是否按照百分比显示
     * @return Bar
     */
    public function barChart(array $data, array $fontStyle = [], bool $showValue = false, bool $showPercentage = false, bool $showSeriesName = false): Bar {
        $barChart = new Bar();
        foreach ($data as $key => $value) {
            $fillType  = $value['attr']['type'] ?? Fill::FILL_NONE;
            $fillColor = $value['attr']['color'] ?? '';
            $series    = new Series($value['name'] ?? '', $value['data']);
            $series->setShowSeriesName($showSeriesName);
            // 柱形填充类型
            if ($fillType != Fill::FILL_NONE) {
                $series->getFill()->setFillType($fillType);
            }
            // 柱形填充色
            if ($fillColor) {
                $series->getFill()->setStartColor(new Color($fillColor));
            }
            $series->setLabelPosition(700);
            // 字体设置
            $font = $series->getFont();
            if (!empty($fontStyle['name'])) {
                $font->setName($fontStyle['name']);
            }
            if (!empty($fontStyle['size'])) {
                $font->setSize($fontStyle['size']);
            }
            if (!empty($fontStyle['color'])) {
                $font->getColor()->setRGB($fontStyle['color']);
            }
            if (isset($fontStyle['bold'])) {
                $font->setBold($fontStyle['bold']);
            }
            $series->setShowValue($showValue);
            if ($showPercentage) {
                $series->setShowPercentage(false);
                $series->setDlblNumFormat('#%');
            }
            $barChart->addSeries($series);
        }
        return $barChart;
    }
}
