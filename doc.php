<?php
$docx = new ZipArchive();

if ($docx->open('file.docx') === true) {	
	$xml = $docx->getFromName('word/document.xml');

	$xml = str_replace('#1', $_POST['firstname'], $xml);
	/*
	$xml = str_replace('#2', '34', $xml);
	$xml = str_replace('#3', '200', $xml);
	$xml = str_replace('#4', '40', $xml);
	$xml = str_replace('#5', '5', $xml);
	*/
	
	echo $xml;
	
	$docx->addFromString('word/document.xml', $xml);	
	
	$docx->close();
}

?>