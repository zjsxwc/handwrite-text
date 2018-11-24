<?php


require_once __DIR__ . '/vendor/autoload.php';


function getFontName()
{
    $fontNames = [
        "HYÑWåtH",
        "liguofu",
        "quietsky"
    ];
    return $fontNames[rand(0, count($fontNames) - 1)];
}

function getFontSize()
{
    $sizeNames = ["21", "21.5", "22", "22.5", "23"];
    return $sizeNames[rand(0, count($sizeNames) - 1)];
}

function getParagraphSpace()
{
    $paragraphSpaces = ["12", "13", "20", "7", "14"];
    return $paragraphSpaces[rand(0, count($paragraphSpaces) - 1)];
}


$text = file_get_contents(__DIR__ . "/text.txt");


// Creating the new document...
$phpWord = new \PhpOffice\PhpWord\PhpWord();

/* Note: any element you append to a document must reside inside of a Section. */

// Adding an empty Section to the document...
$section = $phpWord->addSection();
$textrun = $section->addTextRun();
$textrun->setParagraphStyle(array('spaceAfter' => rand(100,105)));
// Adding Text element with font customized using explicitly created font style object...
$i = 0;
while ($i < mb_strlen($text)) {
    $char = mb_substr($text, $i, 1);
    $i++;
    if ($char == "\n") {
        $section->addTextBreak();
        $textrun = $section->addTextRun();
        $textrun->setParagraphStyle(array('spaceAfter' => rand(100,105)));
        continue;
    }

    $myTextElement = $textrun->addText($char);
    $fontStyle = new \PhpOffice\PhpWord\Style\Font();
    $fontStyle->setName(getFontName());
    $fontStyle->setSize(getFontSize());
    $fontStyle->setPosition(rand(1,3));
    $myTextElement->setFontStyle($fontStyle);
}

// Saving the document as OOXML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
$objWriter->save('output.docx');

// Saving the document as ODF file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'ODText');
$objWriter->save('output.odt');

// Saving the document as HTML file...
$objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'HTML');
$objWriter->save('output.html');
