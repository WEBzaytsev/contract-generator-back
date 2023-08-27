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
}
