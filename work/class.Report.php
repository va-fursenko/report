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
        '21 CENTURY',
        'ASET SG',
        'AUTOM. CONT.',
        'BANK OFFICE',
        'BIRTHDAY',
        'BOOK CLUB',
        'BUSINES MAIL',
        'CERT DB MAIL',
        'CONSULTANTS',
        'CONTIN.GRAD.',
        'COURIER',
        'CROSSSELLING',
        'D-T-D',
        'DB MAIL',
        'DOOR-TO-DOOR',
        'EXHIBITION',
        'EXT DB',
        'FOLLOW UP',
        'GOOGLE ADW',
        'INFO-LINE',
        'INTERNET',
        'KIOSK',
        'LCCIEB',
        'LETTERS',
        'MIR KNIGI',
        'ORIFLAME',
        'PHONE',
        'POST OFFICE',
        'PRINT MEDIA',
        'Phone',
        'Print Media',
        'READERS DIG.',
        'REMAIL',
        'RW STATIONS',
        'STUD BY STUD',
        'TEST PHONE',
        'Телеобзвон',
        'REMAIL 1,2,3',
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

    /** @property array Рабочие данные */
    public $cells = [];

    /** @property array Результат отчёта */
    public $result = [];







    /**
     * Конструктор класса, принимает на вход две матрицы входных данных - 'широкую' и 'узкую' соответственно
     * @param array $firstM
     * @param array $secondM
     * @throws BaseException
     */
    public function __construct($firstM, $secondM)
    {
        // За исключением двух строк в конце первого массива, они должны быть идентичными по ключам

        // Сравниваем длину с учётом постоянных элементов первой матрицы
        if (count($firstM) - count(self::PERMANENT_ROWS) !== count($secondM)){
            throw new BaseException('Размеры файлов не совпадают - ' . count($firstM) . ' (в т.ч. ' . count(self::PERMANENT_ROWS) . ' постоянные) и ' . count($secondM));
        }

        // В первом массиве должны быть постоянные элементы
        if (!Filter::arrayKeyExists(self::PERMANENT_ROWS, $firstM)){
            throw new BaseException("Первый файл не имеет всех постоянных строк ['" . implode("', '", self::PERMANENT_ROWS) . "']'");
        }

        // Пройдём по второму массиву и сольём его со первым, проверяя, чтобы порядок ключей был одинаковым
        reset($firstM);
        $rowIndex = 0; // Счётчик нужен только для того, чтобы вывести его в ошибке
        foreach ($secondM as $keySecond => $value){
            $rowIndex++;
            $keyFirst = each($firstM)['key']; // Ключ первого массива на такой же позиции. Можно было его и засунуть внутрь ряда в отдельный столбец, но можно и так

            // Если на данной позиции ключи не равны, прерываем операцию и приунываем
            if ($keyFirst !== $keySecond){
                throw new BaseException("Файлы расходятся. Ряд $rowIndex, ключи ['$keyFirst'] и ['$keySecond']");
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





    /**
     * Поэтапное почти что ручное вычисление отчёта
     * @return array Результат отчёта - $this->result
     */
    public function process(){
        // Сегодня 15 марта 2016 года, 17:00 мск. Меня зовут Виктор и я начал вручную собирать примерно 384 ряда по 4 ячейки
        $this->result = [];

        // 1.
        $this->result['Print media adverts'] = [
            'prev_SG/TL' => self::sumCells('PRINT MEDIA', [0, 1], ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010', 'PRINT MEDIA  LABIRINT 09.15', 'PRINT MEDIA  LISA KROS.06.15', 'PRINT MEDIA  YEAR 2010', 'PRINT MEDIA  YEAR 2011', 'PRINT MEDIA  Z.RECEPTI 08.14', 'PRINT MEDIA  ZDOROVO']),
            'prev_APP'   => self::sumCells('PRINT MEDIA', [12, 13, 18, 19], ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010', 'PRINT MEDIA  LABIRINT 09.15', 'PRINT MEDIA  LISA KROS.06.15', 'PRINT MEDIA  ZDOROVO', 'PRINT MEDIA  Z.RECEPTI 08.14', 'PRINT MEDIA  YEAR 2011', 'PRINT MEDIA  YEAR 2010']),
            'SG/TL'      => self::sumCells('PRINT MEDIA', [0], ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010']),
            'APP'        => self::sumCells('PRINT MEDIA', [12], ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010']) +
                            self::sumCells('PRINT MEDIA', [18], ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010', 'PRINT MEDIA  LABIRINT 09.15', 'PRINT MEDIA  LISA KROS.06.15', 'PRINT MEDIA  ZDOROVO', 'PRINT MEDIA  Z.RECEPTI 08.14', 'PRINT MEDIA  YEAR 2011'])
        ];
/*
        $this->result[''] = [
            'prev_SG/TL' => self::sumCells('', [], []),
            'prev_APP'   => self::sumCells('', [], []),
            'SG/TL'      => self::sumCells('', [], []),
            'APP'        => self::sumCells('', [], []),
        ];

        $this->result[''] = [
            'prev_SG/TL' => self::sumCells('', [], []),
            'prev_APP'   => self::sumCells('', [], []),
            'SG/TL'      => self::sumCells('', [], []),
            'APP'        => self::sumCells('', [], []),
        ];
*/
        // Да ну нахуй, нереально. Этот метод никогда не будет дописан до конца... Не верю в это
        return $this->result;
    }







    /**
     * Вычисление суммы для одной ячейки результата
     * @param string $groupName Имя группы айдишников
     * @param array $colsIndexes Массив индексов столбцов, которые суммируются
     * @param array $exceptRows Массив айдишников исключений, которые не учитываются
     * @return int
     * @throws BaseException
     */
    public function sumCells($groupName, $colsIndexes, $exceptRows)
    {
        if (!isset($this->rowGroups[$groupName])){
            throw new BaseException("Неизвестная группа айдишников: $groupName (Всего групп " . $this->rowGroupsCount() . ")");
        }

        $sum = 0;

        // Проходим по всем строкам группы, проверяя, что они не в исключениях
        foreach ($this->rowGroups[$groupName] as $rowKey){
            if (!in_array($rowKey, $exceptRows)){

                // Проходим по всем выбранным столбцам
                foreach ($colsIndexes as $col){
                    if (!isset($this->cells[$rowKey][$col])){
                        throw new BaseException("Неверная адресация в массиве: [$rowKey][$col]");
                    }

                    $sum += $this->cells[$rowKey][$col];
                }
            }
        }

        return $sum;
    }







    /** Число строк в рабочем наборе */
    public function rowsCount(){
        return count($this->cells);
    }


    /** Число групп айдишников в рабочем наборе */
    public function rowGroupsCount(){
        return count($this->cells);
    }

}