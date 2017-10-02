# Super simple excel

Simple package for creating an Excel file from arrays when you dont need advanced configuration on cell level.

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

Styles can also be set for headings or a row
```php

```