<?php
/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 16.03.16
 * Time: 9:10
 */

require_once __DIR__ . DIRECTORY_SEPARATOR . 'formulas.php';

/** @const Разделитель столбца и ряда в обозначении ячейки в обработанной формуле */
define('COL_ROW_DELIMITER', '.'); // Точка пока смотрится просто удобнее в дампе

/** @const Разница в индексе ряда между xls и матрицей рабочих данных */
define('ROW_INDEX_DIFF', 4); // Ячейка U8 из xls в нашей матрице имеет индекс [12, 8 - ROW_INDEX_DIFF]. Т.е. сохраняя её формулу, уменьшим индекс ряда на ROW_INDEX_DIFF




/**
 * Преобразование массива формул в удобоваримый для нас вид
 * @param array $arr
 * @return array
 * @throws BaseException
 */
function refactorFormulas($arr)
{
    // Связь между буквенными индексами
    $COL_INDEXES = array_flip(['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']);
    // Разница между индексами рядов в xls и нашей матрице
    $FIRST_ROW_INDEX = 7;

    // Проходим по всем рядам и столбцам
    foreach ($arr as $rowIndex => $row) {
        foreach ($row as $colIndex => $cell) {
            if ($cell === '') {
                continue;
            }
            // Убираем лишние символы, заменяем запятую (перечисление) на +
            $cell = str_replace('$', '', $cell);
            $cell = str_replace(')', '', $cell);
            $cell = str_replace('SUM(', '', $cell);
            $cell = str_replace(',', '+', $cell);

            // Для начала заменяем диапазоны на перечисления. Параллельно меняем символьное обозначение стобцов на числовое
            preg_match_all('/([A-Z]{1,2})(\d{1,5})\:([A-Z]{0,2})(\d{1,5})/i', $cell, $matches);
            foreach ($matches[0] as $i => $diapason){
                // Для наглядности соберём в одном месте координаты диапазона
                $col1 = $matches[1][$i];
                $row1 = $matches[2][$i];
                $col2 = $matches[3][$i] ? $matches[3][$i] : $col1;
                $row2 = $matches[4][$i];
                if (!isset($COL_INDEXES[$col1]) || !isset($COL_INDEXES[$col2])){
                    //throw new BaseException("Кто-то из столбцов '$col1' и '$col2' нам неизвестен");
                    continue;
                }
                $newDiapason = '';
                for ($col = $COL_INDEXES[$col1]; $col <= $COL_INDEXES[$col2]; $col++){
                    for ($row = $row1; $row <= $row2; $row++){
                        $newDiapason .= $newDiapason ? '+' : '';
                        $newDiapason .= $col . COL_ROW_DELIMITER . ($row - ROW_INDEX_DIFF);
                    }
                }
                $cell = str_replace($diapason, $newDiapason, $cell);
            }

            // В ячейке ещё могли остаться одиночные ячейки с символьным обозначением столбца. Меняем его на числовое
            preg_match_all('/([A-Z]{1,2})(\d{2,5})/i', $cell, $matches);
            foreach ($matches[0] as $i => $cl){
                $col = $matches[1][$i];
                $row = $matches[2][$i];
                if (!isset($COL_INDEXES[$col])){
                    //throw new BaseException("Столбец '$col' нам неизвестен");
                    continue;
                }
                $cell = str_replace($matches[0][$i], $COL_INDEXES[$col] . COL_ROW_DELIMITER . ($row - ROW_INDEX_DIFF), $cell);
            }

            // В конце концов, заменяем исходную ячейку обработанной
            $arr[$rowIndex][$colIndex] = $cell;
        }
    }
    return $arr;
}