# Super simple excel

Simple package for creating an Excel file from arrays, traversable objects or Eloquent Collections when you dont need advanced configuration on cell level.

## Installation

This package can be installed via Composer:

``` bash
composer require dmoen/super-simple-excel
```

## Usage

Basic example

```php
$writer = ExcelWriter::create()
    ->setHeadings(["Lorem", "Ipsum", "Sit", "Amet"]);

$writer->addContent(["Row1", "Row1", "Row1", "Row1"]); 
    
$writer->addContent(["Row2", "Row2", "Row2", "Row2"])
    ->save("filepath");    
```

Also work with Eloquent Collections

```php
$writer = ExcelWriter::create()
    ->setHeadings(["Lorem", "Ipsum", "Sit", "Amet"]);
    ->addContent(App\User::all())
    ->save("filepath");    
```

If you want to output the file to the browser:

```php
$writer = ExcelWriter::create()
    ->setHeadings(["Lorem", "Ipsum", "Sit", "Amet"]);

$writer->addContent(["Row1", "Row1", "Row1", "Row1"]); 
    
$writer->addContent(["Row2", "Row2", "Row2", "Row2"])
    ->output("filepath");    
```

Multi dimensional array

```php
ExcelWriter::create()
    ->setHeadings(["Lorem", "Ipsum", "Sit", "Amet"]);
    ->addContent([
        ["Row1", "Row1", "Row1", "Row1"],
        ["Row2", "Row2", "Row2", "Row2"]
    ])
    ->save("filepath");
```    

Default styles can be set for font weight, font size, font type and alignment

```php  
ExcelWriter::create(["bold" => true, "font" => "Arial", "size" => 20, "align" => "center"]);    
```

Styles can also be set for headings or a specific row

```php
$writer = ExcelWriter::create();

$writer->setHeadings(
    ["Lorem", "Ipsum", "Sit", "Amet"],
    [
        "align" => "right",
        "bold"  => false,
        "size" => 15,
        "font" => "Arial"
    ]);
    
$writer->addContent(
    [
        ["Dolore", "Ipsum", "Amet", "Sit"],
        ["Dolore", "Ipsum", "Amet", "Sit"]
    ],
    [
        "align" => "right",
        "bold"  => false,
        "size" => 15,
        "font" => "Arial"
    ]
);
```

Sometimes you want some row spaces between the headings and the content rows:

```php
$writer = ExcelWriter::create();

$writer->setHeadings(
    ["Lorem", "Ipsum", "Sit", "Amet"], [], 1);
```