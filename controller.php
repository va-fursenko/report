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




try {

    switch ($act) {

        // Первая матрица (большая)
        case 'openFirst':
            $result = readMatrix(XLS_FIRST, MATR_FIRST_COLS);
            if ($result['success']) {
                $result['nextStep'] = 'openSecond';
                $result['nextMessage'] = "# Импорт узкой матрицы";
            }
            break;

        // Вторая матрица (маленькая)
        case 'openSecond':
            $result = readMatrix(XLS_SECOND, MATR_SECOND_COLS);
            if ($result['success']) {
                $result['nextStep'] = 'merge';
                $result['nextMessage'] = "# Слияние массивов в общий список";
            }
            break;

        // Сливаем оба полученных массива вместе, получая список новых элементов
        case 'merge':
            $result = mergeMatrixes();
            if ($result['success']) {
                $result['nextStep'] = 'openTemplate';
                $result['nextMessage'] = "# Импорт шаблона";
            }
            break;

        // Сливаем оба полученных массива вместе, получая список новых элементов
        case 'openTemplate':
            $result = readTemplate();
            if ($result['success']) {
                $result['nextStep'] = 'process';
                $result['nextMessage'] = "# Обработка данных";
            }
            break;

        // Сливаем оба полученных массива вместе, получая список новых элементов
        case 'process':
            $result = processData();
            if ($result['success']) {
                //$result['nextStep'] = '';
                //$result['nextMessage'] = "# ";
            }
            break;

        // O_o
        default:
            $result = [
                'success' => false,
                'message' => "Неизвестная команда $act",
            ];
    }

// А что, а вдруг?
}catch(Exception $e){
    Log::save(Log::dumpException($e), CONFIG::ERROR_LOG_FILE);
    $result = [
        'success'   => false,
        'message'   => $e->getMessage(),
    ];
}


// Возвращаем результат
echo json_encode($result);

