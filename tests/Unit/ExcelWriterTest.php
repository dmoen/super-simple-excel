<?php

use PHPUnit\Framework\TestCase;
use Dmoen\SuperSimpleExcel\ExcelWriter;

class ExcelWriterTest extends TestCase
{
    private $samplePath;

    public function __construct($name = null, array $data = [], $dataName = '')
    {
        parent::__construct($name, $data, $dataName);

        $this->samplePath = dirname(dirname(__FILE__))."/excel-sample/sample.xlsx";
    }


    private function readExcel($filename)
    {
        $type = \PHPExcel_IOFactory::identify($filename);
        $reader = \PHPExcel_IOFactory::createReader($type);
        $reader->setReadDataOnly(true);
        $excel = $reader->load($filename);

        $sheet = $excel->getSheet();

        return $sheet->toArray(null, false, false, false);
    }

    public function test_it_creates_headers_correctly()
    {
        $writer = new ExcelWriter();

        $writer->setHeadings(["Lorem", "Ipsum", "Sit", "Amet"], null, 2);
        $writer->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Lorem", "Ipsum", "Sit", "Amet"]
        ], $generatedFile);
    }

    public function test_it_creates_content_correctly()
    {
        $writer = new ExcelWriter();

        $writer->addContent(["Lorem", "Ipsum", "Sit", "Amet"])
            ->addContent(["Dolore", "Ipsum", "Amet", "Sit"])
            ->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Lorem", "Ipsum", "Sit", "Amet"],
            ["Dolore", "Ipsum", "Amet", "Sit"]
        ], $generatedFile);
    }

    public function test_it_creates_multi_dim_arr_content_correctly()
    {
        $writer = new ExcelWriter();

        $writer->addContent([
                ["Lorem", "Ipsum", "Sit", "Amet"],
                ["Dolore", "Ipsum", "Amet", "Sit"]
            ])
            ->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Lorem", "Ipsum", "Sit", "Amet"],
            ["Dolore", "Ipsum", "Amet", "Sit"]
        ], $generatedFile);
    }

    public function test_it_creates_content_and_headers_correctly()
    {
        $writer = new ExcelWriter();

        $writer->setHeadings(["Lorem", "Ipsum", "Sit", "Amet"])
            ->addContent(["Dolore", "Ipsum", "Amet", "Sit"])
            ->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Lorem", "Ipsum", "Sit", "Amet"],
            ["Dolore", "Ipsum", "Amet", "Sit"]
        ], $generatedFile);
    }

    public function test_it_allows_for_header_spacing()
    {
        $writer = ExcelWriter::create(["bold" => true, "font" => "Arial", "size" => 20]);

        $writer->setHeadings(
            ["Lorem", "Ipsum", "Sit", "Amet"],
            [
                "align" => "center",
                "size" => 14
            ], 1)
            ->addContent(
                ["Dolore", "Ipsum", "Amet", "Sit"],
                [
                    "align" => "right",
                    "bold"  => false,
                    "size" => 15
                ]
            )
            ->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Lorem", "Ipsum", "Sit", "Amet"],
            [null, null, null, null],
            ["Dolore", "Ipsum", "Amet", "Sit"]
        ], $generatedFile);
    }
}