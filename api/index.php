<?php
ini_set('default_charset', 'UTF-8');

error_reporting(E_ALL & ~E_DEPRECATED);

require_once '../vendor/autoload.php';
require_once '../src/WordGenerator.php';

$templateFile = '../contract-template.docx';
$array_to_return = array(
    'success' => true,
);

$data = json_decode($_POST['contract'], true);
$static_fields = $data['staticFields'];

$static_fields['clientInitials'] = get_initials($static_fields['clientFirstName'], $static_fields['clientSecondName'], $static_fields['clientLastName']);
$static_fields['executorInitials'] = get_initials($static_fields['executorFirstName'], $static_fields['executorSecondName'], $static_fields['executorLastName']);

$collections = $data['collections'];
$contract_number = $static_fields['contractNumber'];
//$task_content = $data['contractTasks'];

try {
    $wordGenerator = new WordGenerator($templateFile);

    $wordGenerator->setValues($static_fields);
    $wordGenerator->setCollections($collections);
//    $wordGenerator->setTaskContent($task_content);

    $wordGenerator->setOutputFile($contract_number);
    $wordGenerator->save();

    $array_to_return['file_name'] = $wordGenerator->getOutputFile();
} catch (\PhpOffice\PhpWord\Exception\CopyFileException|\PhpOffice\PhpWord\Exception\CreateTemporaryFileException $e) {

    $array_to_return['err_message'] = $e->getMessage();
}

echo json_encode($array_to_return, 1);
die();

function get_initials($firstName = '', $secondName = '', $lastName = ''): string
{
    return $secondName . ' ' . strtoupper(mb_substr($firstName, 0, 1)) . '. ' . strtoupper(mb_substr($lastName, 0, 1)) . '.';
}