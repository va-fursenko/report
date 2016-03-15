<?php

    // Конфиг
    require_once(__DIR__ . '/config.php');
    require_once(__DIR__ . '/lib/inc.common.php');

    // Рисуем страницу
    require_once(CONFIG::ROOT . DIRECTORY_SEPARATOR . CONFIG::TPL_DIR . '/layout.main.php');


$filename = __DIR__ . '/data/sheet2.xml';
$xml = $xml = simplexml_load_file($filename);

$rowCounter = 1;
$result = [];
foreach ($xml->sheetData->row as $row) {
    if ($rowCounter > 7) {
        $counter = 0;
        foreach ($row as $c) {
            if ($counter > 6) {
                break;
            }
            $counter++;
            $attr = $c->attributes();
            $ch = strval($attr['r'][0]);
            if (in_array($ch{0}, ['B', 'C', 'E', 'F'])) {
                $result[$rowCounter][$ch] = [
                    $rowCounter,
                    $ch,
                    isset($c->f) ? (string)$c->f : ''
                ];
            }
        }
    }
    $rowCounter++;
    if ($rowCounter > 390){
        break;
    }
}

function showResult($arr){
    $result = "[\n";
    foreach ($arr as $k => $subArr){
        $result .= "\t$k => [";
        foreach ($subArr as $key => $line) {
            $result .= "'$key' => ['" . implode("', '", $line) . "'], ";
        }
        $result .= "],\n";
    }
    return "$result\n]";
}


echo "<pre style='font-size: 7pt;'>";
echo showResult($result);
echo "</pre>";
