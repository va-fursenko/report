<?php

/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 15.03.16
 * Time: 14:55
 */
class Report
{
    /** @const Постоянные строки в первом файле. Расположены в конце и не участвуют в сравнении */
    const PERMANENT_ROWS = [
        'Телеобзвон',
        'REMAIL 1,2,3'
    ];



    /** @const Группы айдишников, на которые забиваются все строки */
    const GROUP_KEYS = [
        '21 CENTURY',         'ASET SG',            'AUTOM. CONT.',       'BANK OFFICE',
        'BIRTHDAY',           'BOOK CLUB',          'BUSINES MAIL',       'CERT DB MAIL',
        'CONSULTANTS',        'CONTIN.GRAD.',       'COURIER',            'CROSSSELLING',
        'D-T-D',              'DB MAIL',            'DOOR-TO-DOOR',       'EXHIBITION',
        'EXT DB',             'FOLLOW UP',          'GOOGLE ADW',         'INFO-LINE',
        'INTERNET',           'KIOSK',              'LCCIEB',             'LETTERS',
        'MIR KNIGI',          'ORIFLAME',           'PHONE',              'POST OFFICE',
        'PRINT MEDIA',        'Phone',              'Print Media',        'READERS DIG.',
        'REMAIL',             'RW STATIONS',        'STUD BY STUD',       'TEST PHONE',
        'Телеобзвон',         'REMAIL 1,2,3',
    ];



    /** @property array $group Сгруппированные по айдишникам строки
     * [
     *     'INTERNET' => [
     *         'INTERNET'  => [0, 0, 1,...],
     *         'INTERNET2' => [...],
     *         'INTERNET9' => [...],
     *         ...
     *     ],
     *     ...
     * ]
     */
    public $rowGroups = [];
    /** @property array Последняя ошибка */
    public $lastError = '';
    /** @property array Рабочие данные */
    public $cells = [];
    /** @property Список групп айдишников */







    /**
     * Конструктор класса, принимает на вход две матрицы входных данных - 'широкую' и 'узкую' соответственно
     * @param array $firstM
     * @param array $secondM
     * @throws Exception
     */
    public function __construct($firstM, $secondM)
    {
        // За исключением двух строк в конце первого массива, они должны быть идентичными по ключам

        // Сравниваем длину с учётом постоянных элементов первой матрицы
        if (count($firstM) - count(self::PERMANENT_ROWS) !== count($secondM)){
            throw new Exception('Размеры файлов не совпадают - ' . count($firstM) . ' (в т.ч. ' . count(self::PERMANENT_ROWS) . ' постоянные) и ' . count($secondM));
        }

        // В первом массиве должны быть постоянные элементы
        if (!Filter::arrayKeyExists(self::PERMANENT_ROWS, $firstM)){
            throw new Exception("Первый файл не имеет всех постоянных строк ['" . implode("', '", self::PERMANENT_ROWS) . "']'");
        }

        // Пройдём по второму массиву и сольём его со первым, проверяя, чтобы порядок ключей был одинаковым
        reset($firstM);
        $rowIndex = 0; // Счётчик нужен только для того, чтобы вывести его в ошибке
        foreach ($secondM as $keySecond => $value){
            $rowIndex++;
            $keyFirst = each($firstM)['key']; // Ключ первого массива на такой же позиции. Можно было его и засунуть внутрь ряда в отдельный столбец, но можно и так

            // Если на данной позиции ключи не равны, прерываем операцию и приунываем
            if ($keyFirst !== $keySecond){
                throw new Exception("Файлы расходятся. Ряд $rowIndex, ключи ['$keyFirst'] и ['$keySecond']");
            }

            // Соединяем текущий ряд из первой и вторй матрицы
            $this->cells[$keySecond] = array_merge($firstM[$keyFirst], $secondM[$keyFirst]);
        }

        // Дописываем фиксированные строки в конце первой матрицы нулями
        foreach (self::PERMANENT_ROWS as $row){
            $this->cells[$row] = array_pad($firstM[$row], MATR_FIRST_COLS + MATR_SECOND_COLS, 0);
        }

        // Проходим по матрице и раскладываем строки по группам айдишников
        foreach ($this->cells as $key => $row){
            // Подбираем группу, с которой начинается айдишник $key, и прерываем внутренний цикл
            foreach (self::GROUP_KEYS as $groupKey) {
                if (strpos($key, $groupKey) === 0){
                    $this->rowGroups[$groupKey][] = $key; //$groups[$groupKey][$key] = $row;
                    break;
                }
            }
        }

    }




    public function process(){

    }





    /** Число строк в рабочем наборе */
    public function rowsCount(){
        return count($this->cells);
    }

}