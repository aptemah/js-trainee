<?php

//header("Access-Control-Allow-Origin: *");
//header("Content-Type: application/json; charset=UTF-8");



//echo '{"a":1,"b":2,"c":3,"d":4,"e":5}';
$object = $_POST["object"];
$fileName = $_POST["fileName"];

$fp = fopen($fileName, "w");
fwrite($fp, $object);
fclose($fp);
