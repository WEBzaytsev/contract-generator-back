<?php

use PhpOffice\PhpWord\Element\TextRun;
use PhpOffice\PhpWord\Exception\CopyFileException;
use PhpOffice\PhpWord\Exception\CreateTemporaryFileException;
use PhpOffice\PhpWord\TemplateProcessor;
use PhpOffice\PhpWord\PhpWord;
use PhpOffice\PhpWord\IOFactory;

class WordGenerator {
    private $templateFile;
    private $outputFile;
    private $document;

    /**
     * @throws CopyFileException
     * @throws CreateTemporaryFileException
     */
    public function __construct($templateFile) {
        $this->templateFile = $templateFile;

        // Initialize PhpWord TemplateProcessor
        $this->document = new TemplateProcessor($templateFile);
    }

    public function setValue($placeholder, $value) {
        $this->document->setValue($placeholder, $value);
    }

    public function setValues($fields) {
        $this->document->setValues($fields);
    }

    public function setOutputFile($number) {
        $this->outputFile = 'documents/contract_' . $number . '.docx';
    }

    public function setCollections($collections) {

        foreach ($collections as $collection_name => $collection) {
            $collectionWithId = $collection;

            for ($i = 0; $i < count($collectionWithId); $i++) {
                $collectionWithId[$i][$collection_name . 'Id'] = $i + 1;
            }

            $this->document->cloneRowAndSetValues($collection_name . 'Id', $collectionWithId);
        }
    }

    public function save() {
        $this->document->saveAs('../' . $this->outputFile);
    }

    public function getOutputFile() {
        return $this->outputFile;
    }

    /* public function setTaskContent($task_content) {
        $document_with_table = new PhpWord();
        $section = $document_with_table->addSection();
        $section->addText(
            '"Great achievement is usually born of great sacrifice, '
            . 'and is never the result of selfishness." '
            . '(Napoleon Hill)',
            array('bold' => true, 'size' => 20)
        );

        $dom = new DOMDocument();
        $dom->loadHTML($task_content);

        $paragraphs = $dom->getElementsByTagName('p');

        foreach ($paragraphs as $paragraph) {

            $styles = array();

            foreach ($paragraph->childNodes as $node) {
                if ($node instanceof DOMElement) {
                    switch ($node->tagName) {
                        case 'strong':
                            $styles['bold'] = true;
                            break;
                        case 'em':
                            $styles['italic'] = true;
                            break;
                        case 'span':
                            if ($node->hasAttribute('style')) {
                                preg_match('/font-size:\s*(\d+)pt/i', $node->getAttribute('style'), $matches);
                                if (isset($matches[1])) {
                                    $styles['size'] = $matches[1];
                                }
                            }
                            break;
                    }
                }
            }
        }

        $num = rand(1, 100);
        $tmp_name = '../tmp/tmp_' . $num . '.docx';
        $objWriter = IOFactory::createWriter($document_with_table, 'Word2007');
        $objWriter->save($tmp_name);
        $objReader = IOFactory::createReader("Word2007");
        $phpWord = $objReader->load($tmp_name);
        $first_sections = $phpWord->getSection(0);
        $content = '';

        foreach($first_sections->getElements() as $element) {
            if (method_exists($element, 'getElements')) {
                foreach($element->getElements() as $childElement) {
                    if (method_exists($childElement, 'getText')) {
                        $content .= $childElement->getText() . ' ';
                    }
                    else if (method_exists($childElement, 'getContent')) {
                        $content .= $childElement->getContent() . ' ';
                    }
                }
            }
            else if (method_exists($element, 'getText')) {
                $content .= $element->getText() . ' ';
            }
        }
        return $content;

        $fullxml = $objWriter->getWriterPart('Document')->write(); */
//        $tablexml = preg_replace('/^[\s\S]*(<w:sectPr\b.*<\/w:sectPr>).*/', '$1', $fullxml);

        /* return $tablexml;
    } */

    /*public function setTaskContent($task_content) {
        $inline = new TextRun();

        $dom = new DOMDocument();
        $dom->loadHTML($task_content);

        $paragraphs = $dom->getElementsByTagName('p');

        foreach ($paragraphs as $paragraph) {
            $text = $paragraph->textContent;

            $styles = array();

            foreach ($paragraph->childNodes as $node) {
                if ($node instanceof DOMElement) {
                    switch ($node->tagName) {
                        case 'strong':
                            $styles['bold'] = true;
                            break;
                        case 'em':
                            $styles['italic'] = true;
                            break;
                        case 'span':
                            if ($node->hasAttribute('style')) {
                                preg_match('/font-size:\s*(\d+)pt/i', $node->getAttribute('style'), $matches);
                                if (isset($matches[1])) {
                                    $styles['size'] = $matches[1];
                                }
                            }
                            break;
                    }
                }
            }

            $inline->addText($text, $styles);
//            $inline->addTextBreak(2);
        }

        $this->document->setComplexValue('contractTasks', $inline);
    }*/
}
