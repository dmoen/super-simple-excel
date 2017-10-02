<?php

use PHPUnit\Framework\TestCase;
use Dmoen\SuperSimpleExcel\ExcelWriter;
use Illuminate\Support\Collection;
use Illuminate\Database\Eloquent\Model;

class TestModel extends Model{
    protected $guarded = [];
};

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

        $writer->setHeadings(["Lorem", "Ipsum", "Sit", "Amet"], [], 2);
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
                [
                    ["Dolore", "Ipsum", "Amet", "Sit"]
                ],
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

    public function test_it_works_with_collections()
    {
        $collection = new class implements Iterator {

            private $var1 = [["Lorem", "Ipsum", "Sit", "Amet"], ["Dolore", "Ipsum", "Amet", "Sit"]];

            private $key = 0;

            public function current()
            {
                return $this->var1[$this->key];
            }

            public function next()
            {
                return $this->var1[$this->key++];
            }

            public function key()
            {
                return $this->key;
            }

            public function valid()
            {
                return $this->key < 2;
            }

            public function rewind(){}
        };

        $writer = new ExcelWriter();

        $writer->addContent($collection)
            ->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Lorem", "Ipsum", "Sit", "Amet"],
            ["Dolore", "Ipsum", "Amet", "Sit"]
        ], $generatedFile);
    }

    public function test_it_works_with_nested_collections()
    {
        $collection = new class implements Iterator {

            private $var1 = ["Lorem", "Ipsum", "Sit", "Amet"];

            private $key = 0;

            public function current()
            {
                return $this->var1[$this->key];
            }

            public function next()
            {
                return $this->var1[$this->key++];
            }

            public function key()
            {
                return $this->key;
            }

            public function valid()
            {
                return $this->key < 4;
            }

            public function rewind(){}
        };

        $writer = new ExcelWriter();

        $writer->addContent([$collection, clone $collection])
            ->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Lorem", "Ipsum", "Sit", "Amet"],
            ["Lorem", "Ipsum", "Sit", "Amet"]
        ], $generatedFile);
    }

    public function test_it_works_with_laravel_collections_with_models()
    {
        $model1 = new TestModel([
            "user" => "Bill",
            "role" => "Admin",
            "city" => 'Stockholm',
            "car" => "Audi"
        ]);

        $model2 = new TestModel([
            "use2" => "George",
            "role" => "User",
            "city" => 'Gothenburg',
            "car" => "Volvo"
        ]);

        $collection = new Collection([$model1, $model2]);

        $writer = new ExcelWriter();

        $writer->addContent($collection)
            ->save($this->samplePath);

        $generatedFile = $this->readExcel($this->samplePath);

        $this->assertSame([
            ["Bill", "Admin", 'Stockholm', "Audi"],
            ["George", "User", 'Gothenburg', "Volvo"]
        ], $generatedFile);
    }
}