<?php
/**
 * Импорт xls и всё с этим связанное
 * User: viktor
 * Date: 10.03.16
 * Time: 14:11
 */


// Папка для рабочих данных
define('XLS_ROOT', CONFIG::ROOT . DIRECTORY_SEPARATOR . 'data/');
// Расширение входных файлов
define('XLS_EXT', '.xls');
// Имя временного файла для двух объединённых матриц
define('XLS_MERGED', 'XLS_MERGED');

// Исходные xls файлы
define('XLS_FIRST', 'WEEKC_CH');
define('XLS_SECOND', 'reclama3');
define('XLS_TEMPLATE', 'Template(month)new');






// Индекс страницы в файле отчёта
define('TPL_SHEET', 0);
// Стартовая строка, с которой начинается полезная информация в файле отчёта
define('TPL_FIRST_ROW', 3);
// Стартовый столбец, с которого начинается полезная информация в файле отчёта
define('TPL_FIRST_COL', 0);
// Ширина первой (большой) матрицы
define('MATR_FIRST_COLS', 18);
// Ширина второй (маленькой) матрицы
define('MATR_SECOND_COLS', 6);






/**
 * Mr Hankey's christmas classics
 * @param array $a
 * @return string
 */
function showArr($a){
    $result = '';
    foreach ($a as $k => $v){
        $result .= "\t" . Filter::strPad("'$k'", 32) . " => [" . implode(', ', $v) . "]\n";
    }
    return "[\n$result]";
}






/**
 * Читаем из указанного файла с указанной матрицы
 * @param string $filename  Имя файла
 * @param int $sheetN       Номер страницы в файле
 * @param int $colStartFrom Столбец, с которого начинается чтение
 * @param int $rowStartFrom Ряд, с которого начинается чтение
 * @param int $colCount     Ширина считываемой матрицы
 * @return mixed
 */
function getMatrix($filename, $sheetN, $colStartFrom, $rowStartFrom, $colCount){
    // Открываем файл, устанавливаем индекс активного листа и получаем его
    $xls = PHPExcel_IOFactory::load(XLS_ROOT .$filename . XLS_EXT);
    $xls->setActiveSheetIndex($sheetN);
    $sheet = $xls->getActiveSheet();

    // Строка, с которой начинаем сбор айдишников в столбце
    $rowCounter = $rowStartFrom;
    $matrix = [];

    // Смотрим, что в первой рабочей строке
    $val = trim($sheet->getCellByColumnAndRow($colStartFrom, $rowCounter)->getValue());
    while ($val !== '' && $val !== 'Total') { // Грубовато, но заканчиваем чтение стобца

        // Добавляем строку с указанным айдишником
        $matrix[$val] = [];
        // Собираем матрицу указанной ширины
        for ($i = 1; $i <= $colCount; $i++){
            $matrix[$val][] = intval($sheet->getCellByColumnAndRow($colStartFrom + $i, $rowCounter)->getValue());
        }

        // Переходим к следующему ряду
        $rowCounter++;
        $val = trim($sheet->getCellByColumnAndRow($colStartFrom, $rowCounter)->getValue());
    }
    return $matrix;
};







/**
 * Читаем первую матрицу (большую)
 * @return mixed
 */
function readFirstMatrix(){
    // Читаем и сериализуем
    $matrix = getMatrix(XLS_FIRST, TPL_SHEET, TPL_FIRST_COL, TPL_FIRST_ROW, MATR_FIRST_COLS);
    file_put_contents(XLS_ROOT . XLS_FIRST, json_encode($matrix));

    return [
        'success' => true,
        'message' => "Считано записей: " . count($matrix)// . "\n" . showArr($matrix)
    ];
}







/**
 * Читаем вторую матрицу (маленькую)
 * @return mixed
 */
function readSecondMatrix(){
    // Читаем и сериализуем
    $matrix = getMatrix(XLS_SECOND, TPL_SHEET, TPL_FIRST_COL, TPL_FIRST_ROW, MATR_SECOND_COLS);
    file_put_contents(XLS_ROOT . XLS_SECOND, json_encode($matrix));

    return [
        'success' => true,
        'message' => "Считано записей: " . count($matrix)// . "\n" . showArr($matrix)
    ];
}








/**
 * Импорт файла шаблона
 * @return array
 */
function readTemplate(){
    $sheetN       = 1;  // Номер страницы в документе. В шаблоне это вторая страница
    $colStartFrom = 6;  // Первый столбец для выборки
    $rowCounter   = 4;  // Первый ряд для выборки - такое ощущение, что счёт почему-то идёт с 1
    $colCount     = [   // Число столбцов для чтения - два фрагмента с заданной ранее шириной
        MATR_FIRST_COLS,
        MATR_SECOND_COLS
    ];

    // Открываем файл, устанавливаем индекс активного листа и получаем его
    $xls = PHPExcel_IOFactory::load(XLS_ROOT . XLS_TEMPLATE . XLS_EXT);
    $xls->setActiveSheetIndex($sheetN);
    $sheet = $xls->getActiveSheet();

    $matrix = [];

    // Смотрим, что в первой рабочей строке
    $val = trim($sheet->getCellByColumnAndRow($colStartFrom, $rowCounter)->getValue());
    while ($val !== '' && $val !== 'Total') { // Грубовато, но заканчиваем чтение стобца

        // Добавляем строку с указанным айдишником
        $matrix[$val] = [];
        // Собираем первую матрицу указанной ширины
        for ($i = 1; $i <= $colCount[0]; $i++){
            $matrix[$val][] = intval($sheet->getCellByColumnAndRow($colStartFrom + 1 + $i, $rowCounter)->getValue());
        }
        // Собираем вторую матрицу указанной ширины
        for ($i = 1; $i <= $colCount[1]; $i++){
            $matrix[$val][] = intval($sheet->getCellByColumnAndRow($colStartFrom + 3 + $colCount[0] + $i, $rowCounter)->getValue());
        }

        // Переходим к следующему ряду
        $rowCounter++;
        $val = trim($sheet->getCellByColumnAndRow($colStartFrom, $rowCounter)->getValue());
    }

    // Сериализуем матрицу шаблона
    file_put_contents(XLS_ROOT . XLS_TEMPLATE, json_encode($matrix));

    return [
        'success' => true,
        'message' => "Считано записей: " . count($matrix) . "\n" . showArr($matrix)
    ];
}

