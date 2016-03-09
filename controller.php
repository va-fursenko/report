<?php
/**
 * Поэтапное получение отчёта
 * User: viktor
 * Date: 09.03.16
 * Time: 10:07
 */

if (!isset($_GET['action'])) {
    exit("Прощай, со всех вокзалов поезда уходят в дальние края. Прощай, мы расстаёмся навсегда под белым небом янваааааряяя!...");
}
$act = $_GET['action'];


// Конфиг
require_once(__DIR__ . '/config.php');
require_once(__DIR__ . '/lib/inc.common.php');

// Рабочий модуль
require_once(__DIR__ . '/work/report.php');



// Исходный экселевский файл
$filename = CONFIG::ROOT . DIRECTORY_SEPARATOR . 'data/Template(month)new.xls';



switch ($act){

    // Открываем файл и читаем список айдишников в столбце
    case 'openXLS':
        $result = makeDaddyHappy($filename);
        $result = [
            'success' => true,
            'message' => $result,
        ];
        break;

    // O_o
    default:
        $result = [
            'success' => false,
            'message' => "Неизвестная команда $act",
        ];
}


// Возвращаем результат
echo json_encode($result);

