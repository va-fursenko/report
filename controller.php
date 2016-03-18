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
require_once(__DIR__ . '/work/import.xls.php');
require_once(__DIR__ . '/work/class.Report.php');




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
                $result['nextMessage'] = '# Слияние массивов в общий список';
            }
            break;





        // Сливаем оба полученных массива вместе, получая список новых элементов
        case 'merge':

            // Пробуем создать объект отчёта

            // Создаём объект отчёта
            $report = new Report();
            $report->merge(
                json_decode(file_get_contents(XLS_ROOT . XLS_FIRST), true),
                json_decode(file_get_contents(XLS_ROOT . XLS_SECOND), true)
            );

            // Сериализуем объект
            file_put_contents(XLS_ROOT . JSON_REPORT_FILE, json_encode($report));

            // Возвертаем осмысленное послание
            $result = [
                'success'       => true,
                'nextStep'      => 'process',
                'nextMessage'   => '# Вычисление отчёта',
                'message'       => "Записей в объединённом массиве: " . $report->rowsCount() //. "\nИтог:" . $report->showCells(false),
            ];


            break;





        // Вычисляем-вычисляем...
        case 'process':

            $report = new Report();
            $report->jsonImport(file_get_contents(XLS_ROOT . JSON_REPORT_FILE));

            $result = $report->process();

            if ($result['success']) {
                $result['userData'] = $report->getResultUserData();
                $result['message']  = '';
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

