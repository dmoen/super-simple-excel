<?php

namespace Dmoen\SuperSimpleExcel;

class ExcelWriter extends \PHPExcel
{
    private $currentRow = 1;

    private $sheetOptions = [
        'font'  => null,
        'size'  => null,
        'align' => null,
        'bold'  => null
    ];

    public function __construct(Array $sheetOptions = null)
    {
        parent::__construct();

        if($sheetOptions){
            $this->sheetOptions = array_merge(
                $this->sheetOptions, $sheetOptions
            );
        }

        $this->setDefaultStyles();
    }

    public static function create(Array $sheetOptions = null)
    {
        return new static($sheetOptions);
    }

    private function setDefaultStyles()
    {
        $this->getDefaultStyle()
            ->applyFromArray($this->buildStylesArr($this->sheetOptions));
    }

    private function buildStylesArr(Array $styleOpts)
    {
        return [
            'font' => [
                'name' => $styleOpts['font'],
                'size' => $styleOpts['size'],
                'bold' => $styleOpts['bold']
            ],
            'alignment' => [
                'horizontal' => $styleOpts['align']
            ]
        ];
    }

    private function numToAlpha($n)
    {
        for($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }

        return $r;
    }

    public function setHeadings(Array $headings, $styles = ["bold" => true], $rowSpacing = 0)
    {
        foreach(array_values($headings) as $index => $heading){
            $cell = $this->numToAlpha($index).$this->currentRow;
            $this->setCell($cell, $heading);
            $this->setCellStyles($cell, $styles);
        }

        $this->currentRow += 1 + $rowSpacing;

        return $this;
    }

    private function setCell($cell, $value)
    {
        preg_match("/^([a-z])+/i", 'SDC1', $matches);

        $col = $matches[0];

        $sheet = $this->GetActiveSheet();

        $sheet->SetCellValue($cell, $value);

        $sheet->GetColumnDimension($col)
            ->setAutoSize(true);

        return $this;
    }

    public function addContent($content, $styles = [])
    {
        if(is_array($content) || $content instanceof \Traversable){
            foreach($content as $index => $rowContent){
                if($index == 0 && !is_array($rowContent) && !$content instanceof \Traversable){
                    $this->setRowContent($this->currentRow++, $content, $styles);
                    return $this;
                }

                $this->setRowContent($this->currentRow++, $rowContent, $styles);
            }
        }

        return $this;
    }

    private function setRowContent($rowNr, $rowContent, $styles)
    {
        $rowContent = method_exists($rowContent,'toArray') ?
            $rowContent->toArray() : $rowContent;
        $colNr = 0;

        foreach($rowContent as $index => $cellValue) {
            $cell = $this->numToAlpha($index).$rowNr;
            $this->setCell($cell, $cellValue);
            $this->setCellStyles($cell, $styles);

            $colNr++;
        }
    }

    private function setCellStyles($cell, $styles)
    {
        $styles = array_merge(
            $this->sheetOptions, $styles
        );

        $this->getActiveSheet()
            ->getStyle($cell)
            ->applyFromArray($this->buildStylesArr($styles));
    }

    public function save($filename)
    {
        return (new \PHPExcel_Writer_Excel2007($this))
            ->save($filename);
    }

    public function output($filename)
    {
        $tmp = tempnam("/tmp", "excel_");

        $this->save($tmp);
        $data = file_get_contents($tmp);

        if (substr($filename, -5) != ".xlsx") {
            $filename .= ".xlsx";
        }

        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
        header("Content-Length: " . strlen($data));

        echo $data;

        unlink($tmp);
    }
}