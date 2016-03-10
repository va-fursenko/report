<?php
/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 04.03.16
 * Time: 14:59
 */


require_once (__DIR__ . DIRECTORY_SEPARATOR . 'import.xls.php');







/**
 * Слияние старого и нового списка айдишников
 * @return mixed
 */
function mergeMatrixes(){
    // Постоянные строки в первом файле. Расположены в конце и не участвуют в сравнении
    $permanentRows = ['Телеобзвон', 'REMAIL 1'];

    // Достаём подготовленные ранее данные
    $firstM = json_decode(file_get_contents(XLS_ROOT . 'XLS_FIRST'), true);
    $secondM = json_decode(file_get_contents(XLS_ROOT . 'XLS_SECOND'), true);

    // За исключением двух строк в конце первого массива, они должны быть идентичными по ключам

    // Для начала сравниваем длину
    if (count($firstM) - count($permanentRows) !== count($secondM)){
        return [
            'success'   => false,
            'message'   => 'Размеры файлов не совпадают - ' . count($firstM) . ' (в т.ч. ' . count($permanentRows) . ' постоянные) и ' . count($secondM),
        ];
    }

    // В первом массиве должны быть постоянные элементы
    if (!Filter::arrayKeyExists($permanentRows, $firstM)){
        return [
            'success'   => false,
            'message'   => "Первый файл не имеет всех постоянных строк ['" . implode("', '", $permanentRows) . "']'",
        ];
    }

    // Пройдём по второму массиву и сольём его со первым, проверяя, чтобы порядок ключей был одинаковым
    reset($firstM);
    $rowIndex = 0; // Счётчик нужен только для того, чтобы вывести его в ошибке
    foreach ($secondM as $key => $value){
        $rowIndex++;
        $firstKey = each($firstM)['key']; // Ключ первого массива на такой же позиции. Можно было его и засунуть внутрь ряда в отдельный столбец, но можно и так

        // Если на данной позиции ключи не равны, прерываем операцию и приунываем
        if ($firstKey !== $key){
            return [
                'success'   => false,
                'message'   => "Файлы расходятся. Ряд $rowIndex, ключи ['$firstKey'] и ['$key'] соответственно",
            ];
        }
        $firstM[$key] = array_merge($firstM[$key], $secondM[$key]);
    }

    // Дописываем фиксированные строки в конце первой матрицы нулями
    foreach ($permanentRows as $row){
        $firstM[$row] = array_pad($firstM[$row], MATR_FIRST_COLS + MATR_SECOND_COLS, 0);
    }

    // Сериализуем объединённую матрицу
    file_put_contents(XLS_ROOT . XLS_MERGED, json_encode($firstM));

    return [
        'success'   => true,
        'message'   => "Записей в объединённом массиве: " . count($firstM) . "\nИтог:\n" . showArr($firstM),
        //    'message'   => Log::printObject($baseCol),
    ];
}








/**
 * Обработка полученных ранее данных
 * @return array
 */
function processData(){
    // Достаём подготовленные ранее данные
    $firstM = json_decode(file_get_contents(XLS_ROOT . 'XLS_FIRST'), true);
    $secondM = json_decode(file_get_contents(XLS_ROOT . 'XLS_SECOND'), true);
    $templateM = json_decode(file_get_contents(XLS_ROOT . 'XLS_TEMPLATE'), true);

    return [
        'success'   => true,
        'message'   => "Считано записей: " . count($templateM),
    ];
}