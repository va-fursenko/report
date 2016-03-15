<?php
/**
 * Импорт xls и всё с этим связанное
 * User: viktor
 * Date: 10.03.16
 * Time: 14:11
 */


// Папка для рабочих данных
define('XLS_ROOT', CONFIG::ROOT . DIRECTORY_SEPARATOR . 'data' . DIRECTORY_SEPARATOR);
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



/** @todo А зачем нам вообще теперь нужно морочиться с экселем, если можно читать данные из цсв? */



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
 * Читаем матрицу из указанного файла по указанному параметру
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
    $val = iconv('cp866', 'utf-8', $val);
    while ($val !== '' && $val !== 'Total') { // Грубовато, но заканчиваем чтение стобца

        // Все айдишники должны быть уникальными
        if (isset($matrix[$val])){
            return $val;
        }

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
 * Читаем первую (широкую), или вторую (узкую) матрицу
 * @param string $fileName Имя файла без расширения
 * @param int $colCount Число столбцов для считывания
 * @return mixed
 */
function readMatrix($fileName, $colCount){
    // Читаем и сериализуем
    $matrix = getMatrix($fileName, TPL_SHEET, TPL_FIRST_COL, TPL_FIRST_ROW, $colCount);

    // Если вместо массива вернулась строка, значит это имя повторяющегося ключа
    if (is_string($matrix)){
        return [
            'success' => false,
            'message' => "Повторяющийся ключ: $matrix",
        ];
    }

    file_put_contents(XLS_ROOT . $fileName, json_encode($matrix));

    return [
        'success' => true,
        'message' => "Считано записей: " . count($matrix)// . "\n" . showArr($matrix)
    ];
}

