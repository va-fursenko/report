<?php

require_once(__DIR__ . DIRECTORY_SEPARATOR . 'config.Report.php');





/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 15.03.16
 * Time: 14:55
 */
class Report
{

    /**
     * @const Разделитель столбца и ряда в обозначении ячейки в обработанной формуле
     * ОДИН СИМВОЛ
     */
    const COL_ROW_DELIMITER = '.'; // Точка пока смотрится просто удобнее в дампе



    /** @const Разница в индексе ряда между xls и матрицей рабочих данных */
    const ROW_INDEX_DIFF = 4; // Ячейка U8 из xls в нашей матрице имеет индекс [12, 8 - ROW_INDEX_DIFF]. Т.е. сохраняя её формулу, уменьшим индекс ряда на ROW_INDEX_DIFF



    /** @property array $group Сгруппированные айдишники
     * [
     *     'INTERNET' => ['INTERNET', 'INTERNET 2', 'INTERNET 9', ... ],
     *     ...
     */
    public $rowGroups = [];




    // Инициализуется в result.captions.php Да, говнокод
    /** @const Константная ячейка */
    public static $CELL_I1 = -1000000; // Чтобы не забыть инициализовать

    /** @const Заголовки рядов результата */
    public static $RESULT_CAPTIONS = [];

    /** @const Группы айдишников, на которые забиваются все строки */
    public static $GROUP_KEYS = [];

    /** @const Постоянные строки в первом файле. Расположены в конце и не участвуют в сравнении */
    public static $PERMANENT_ROWS = [];

    /** @const Связь между буквенной индексацией столбцов в xls и в матрице рабочих данных */
    public static $COL_INDEXES = [];

    /** @property array Рабочие данные */
    protected $_cells = [];

    /** @property array Жёлтые ряды, которые вычисляются по одинаково простой схеме. Прогоним их через цикл */
    public $singleResultRows = [];

    /** @property array Результат отчёта */
    public $result = [];



    /**
     * @property array $keyIndexes
     * Порядковые (с 0) номера айдишников [0 => '21 CENTURY', 1 => 'ASET SG',...]
     * и обратные пары ['21 CENTURY' => 0, 'ASET SG' => 1,...]
     * Короче говоря, двусторонняя ишачья залупа
     */
    protected $keyIndexes = [];



    /**
     * @property array $formulas
     * Тексты формул из оригинального файла отчёта
     * <Шутка про фистинг/>
     *
     * Индексация идёт по ряду xls
     * Если в ячейке есть формула суммирования, то оттуда выдирается строка с ячейками,
     * буквенная индексация столбцов меняется на числовую, а числовая индексакция строк поправляется,
     * чтобы соответствовать индексации в матрице рабочих данных
     * Разделитель координат может быть любым и задаётся
     * Например, '21' => ['E' => '0.123', 'F' => '12.123+18.123']
     */
    public $formulas = [];



    /**
     * Универсальный геттер ячейки рабочих данных
     * @param int|string $row Числовой индекс или строковый ключ ряда
     * @param int $col Индекс столбца
     * @return int
     * @throws BaseException
     */
    public function cell($row, $col)
    {
        if (is_numeric($row)) {
            $row = intval($row);
        }
        if (is_int($row) && isset($this->keyIndexes[$row])) {
            $row = $this->keyIndexes[$row];
        }
        if (!is_numeric($col)){
            $col = self::colIndex($col);
        }
        if (!isset($this->_cells[$row][$col])) {
            throw new BaseException("Неверная адресация в массиве: [$row][$col]");
        }
        return $this->_cells[$row][$col];
    }

    public function cells()
    {
        return $this->_cells;
    }




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
        if (count($firstM) - count(self::$PERMANENT_ROWS) !== count($secondM)){
            throw new BaseException('Размеры файлов не совпадают - ' . count($firstM) . ' (в т.ч. ' . count(self::$PERMANENT_ROWS) . ' постоянные) и ' . count($secondM));
        }

        // В первом массиве должны быть постоянные элементы
        if (!Filter::arrayKeyExists(self::$PERMANENT_ROWS, $firstM)){
            throw new BaseException("Первый файл не имеет всех постоянных строк ['" . implode("', '", self::$PERMANENT_ROWS) . "']'");
        }

        // Пройдём по второму массиву и сольём его со первым, проверяя, чтобы порядок ключей был одинаковым
        reset($firstM);
        $rowIndex = 0;
        foreach ($secondM as $keySecond => $value) {
            $rowIndex++;
            $keyFirst = each($firstM)['key']; // Ключ первого массива на такой же позиции. Можно было его и засунуть внутрь ряда в отдельный столбец, но можно и так

            // Если на данной позиции ключи не равны, прерываем операцию и приунываем
            if ($keyFirst !== $keySecond) {
                throw new BaseException("Файлы расходятся. Ряд $rowIndex, ключи ['$keyFirst'] и ['$keySecond']");
            }

            // Соединяем текущий ряд из первой и второй матрицы
            $this->_cells[$keySecond] = array_merge($firstM[$keyFirst], $secondM[$keyFirst]);
            $this->keyIndexes[$rowIndex - 1] = $keySecond;
            $this->keyIndexes[$keySecond] = $rowIndex - 1;
        }

        // Дописываем фиксированные строки в конце первой матрицы нулями
        foreach (self::$PERMANENT_ROWS as $index => $row) {
            $this->_cells[$row] = array_pad($firstM[$row], MATR_FIRST_COLS + MATR_SECOND_COLS, 0);
            $this->keyIndexes[$rowIndex + $index] = $row;
            $this->keyIndexes[$row] = $rowIndex + $index;
        }

        // Человеку с фамилией Итого всегда достаётся больше других
        $this->_cells['Total'] = array_pad([], MATR_FIRST_COLS + MATR_SECOND_COLS, 0);
        foreach ($this->_cells['Total'] as $col => $value) {
            $this->_cells['Total'][$col] = array_sum(array_column($this->_cells, $col));
        }
        $this->keyIndexes['Total'] = count($this->keyIndexes) - 1;
        $this->keyIndexes[count($this->keyIndexes) - 1] = 'Total';

        // Проходим по матрице и раскладываем строки по группам айдишников
        foreach ($this->_cells as $key => $row) {
            // Подбираем группу, с которой начинается айдишник $key, и прерываем внутренний цикл
            foreach (self::$GROUP_KEYS as $groupKey) {
                if (stripos($key, $groupKey) === 0 && !in_array($key, self::$PERMANENT_ROWS)) {
                    $this->rowGroups[$groupKey][] = $key; //$groups[$groupKey][$key] = $row;
                    break;
                }
            }
        }

        /* Собираем простые ряды результата
         * Для начала соберём все исключения, не участвующие в результате.
         * Это фиксированные ряды, "Итого" и ряды, уже участвовавшие в формулах в рядах выше жёлтых ячеек
         * Исключения берутся только из формул в столбцах prev_SG/TL и prev_APP
         */
        $exceptRows = array_merge(
            self::$PERMANENT_ROWS,
            ['Total'],

            // Исключаем всех, кроме исключений %) Ну не создавать же для одного айдишника массив с белым списком, или хуячить его в сравнение if(...)
            array_diff($this->rowGroups['PRINT MEDIA'], ['PRINT MEDIA  HAPPY_FAMILIY']),

            // Некоторые группы айдишников просчитывались целиком, целиком их и исключаем
            $this->rowGroups['BIRTHDAY'],
            $this->rowGroups['REMAIL'],
            $this->rowGroups['CONTIN.GRAD.'],
            $this->rowGroups['COURIER'],
            $this->rowGroups['CONSULTANTS'],
            $this->rowGroups['STUD BY STUD'],
            $this->rowGroups['INFO-LINE'],
            $this->rowGroups['LETTERS'],

            // Отдельные айдишники-маргиналы
            ['FOLLOW UP', 'INTERNET', 'INTERNET     LETNIE CENI2011', 'INTERNET     LET CENI2011 TM', 'INTERNET     POP-UP WINDOW', 'INTERNET     TARGET.MYMIR']
        );

        // Эту и предыдущую операцию можно было сделать в одном цикле, но мне кажется, что тут важнее хоть какая-то читаемость кода
        foreach ($this->_cells as $key => $row) {
            if (!in_array($key, $exceptRows) && $this->sumCells(['I', 'U', 'AC'], $key) > 0) {
                $this->singleResultRows[] = $key;
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

        #0 PRINT MEDIA
        $this->result['Print media adverts'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],                 ['PRINT MEDIA'],    ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010', 'PRINT MEDIA  LABIRINT 09.15', 'PRINT MEDIA  LISA KROS.06.15', 'PRINT MEDIA  YEAR 2010', 'PRINT MEDIA  YEAR 2011', 'PRINT MEDIA  Z.RECEPTI 08.14', 'PRINT MEDIA  ZDOROVO']),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],     ['PRINT MEDIA'],    ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010', 'PRINT MEDIA  LABIRINT 09.15', 'PRINT MEDIA  LISA KROS.06.15', 'PRINT MEDIA  YEAR 2010', 'PRINT MEDIA  YEAR 2011', 'PRINT MEDIA  Z.RECEPTI 08.14', 'PRINT MEDIA  ZDOROVO']),
            'SG/TL'      => $this->sumCells(['I'],                      ['PRINT MEDIA'],    ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010']),
            'APP'        => $this->sumCells(['U'],                      ['PRINT MEDIA'],    ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010'])
                          + $this->sumCells(['AC'],                     ['PRINT MEDIA'],    ['PRINT MEDIA  HAPPY_FAMILIY', 'PRINT MEDIA  INSERTS', 'PRINT MEDIA  INSERTS 2010', 'PRINT MEDIA  LABIRINT 09.15', 'PRINT MEDIA  LISA KROS.06.15',                           'PRINT MEDIA  YEAR 2011', 'PRINT MEDIA  Z.RECEPTI 08.14', 'PRINT MEDIA  ZDOROVO'])
        ];

        #1
        $this->result['Print media inserts'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],               'PRINT MEDIA  INSERTS'),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],   'PRINT MEDIA  INSERTS'),
            'SG/TL'      => $this->sumCells(['I'],                    'PRINT MEDIA  INSERTS')
                          + $this->sumCells(['I'],                    'PRINT MEDIA  INSERTS 2010'),
            'APP'        => $this->sumCells(['U', 'AC'],              'PRINT MEDIA  INSERTS')
                          + $this->sumCells(['U', 'AC'],              'PRINT MEDIA  INSERTS 2010')
        ];

        #2 REMAIL
        $this->result['Remails'] = [
            'prev_SG/TL' => $this->sumCells((['I', 'J']),             ['REMAIL'])
                          + $this->sumCells((['I', 'J']),             'REMAIL 1,2,3'),
            'prev_APP'   => $this->sumCells((['U', 'V', 'AC', 'AD']), ['REMAIL'])
                          + $this->sumCells((['U', 'V']),             'REMAIL 1,2,3'),
            'SG/TL'      => $this->sumCells(['I'],                    ['REMAIL'])
                          + $this->sumCells(['I'],                    'REMAIL 1,2,3'),
            'APP'        => $this->sumCells(['U', 'AC'],              ['REMAIL'])
                          + $this->sumCells(['U', 'AC'],              'REMAIL 1,2,3')
        ];

        #3
        $this->result['Follow up'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],               'FOLLOW UP'),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],   'FOLLOW UP'),
            'SG/TL'      => $this->sumCells(['I'],                    'FOLLOW UP'),
            'APP'        => $this->sumCells(['U', 'AC'],              'FOLLOW UP')
        ];

        #4 CONTIN.GRAD.
        $this->result['Continuation graduates'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],               ['CONTIN.GRAD.']),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],   ['CONTIN.GRAD.']),
            'SG/TL'      => $this->sumCells(['I'],                    ['CONTIN.GRAD.']),
            'APP'        => $this->sumCells(['U', 'AC'],              ['CONTIN.GRAD.'])
        ];

        #5
        $this->result['Internet'] = [
            'prev_SG/TL' => $this->sumCells(['J'],         'INTERNET')
                          + $this->sumCells(['J'],         'INTERNET     POP-UP WINDOW'),
            'prev_APP'   => $this->sumCells(['V', 'AD'],   'INTERNET')
                          + $this->sumCells(['V', 'AD'],   'INTERNET     POP-UP WINDOW'),
            'SG/TL'      => $this->sumCells(['I'],         'INTERNET')
                          + $this->sumCells(['I'],         'INTERNET     POP-UP WINDOW')
                          + $this->sumCells(['I'],         'INTERNET     LETNIE CENI2011'),
            'APP'        => $this->sumCells(['I', 'AC'],   'INTERNET')
                          + $this->sumCells(['I', 'AC'],   'INTERNET     POP-UP WINDOW')
                          + $this->sumCells(['I', 'AC'],   'INTERNET     LETNIE CENI2011')
        ];

        #6
        $this->result['Internet Load'] = [
            'prev_SG/TL' => $this->sumCells(['J'],         'INTERNET     LOAD'),
            'prev_APP'   => $this->sumCells(['V', 'AD'],   'INTERNET     LOAD'),
            'SG/TL'      => $this->sumCells(['I'],         'INTERNET     LOAD') + self::$CELL_I1,
            'APP'        => $this->sumCells(['U', 'AC'],   'INTERNET     LOAD')
        ];


        #7 Телеобзвон
        $this->result['Telemarketing'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],    'Телеобзвон'),
            'prev_APP'   => $this->sumCells(['U', 'V'],   'Телеобзвон'),
            'SG/TL'      => $this->sumCells(['I'],         'Телеобзвон')
                          + $this->sumCells(['I'],         'INTERNET     LET CENI2011 TM')
                          + $this->sumCells(['I'],         'INTERNET     TARGET.MYMIR'),
            'APP'        => $this->sumCells(['U'],         'Телеобзвон')
                          + $this->sumCells(['U', 'AC'],   'INTERNET     TARGET.MYMIR')
                          + $this->sumCells(['U', 'AC'],   'INTERNET     LET CENI2011 TM')
        ];

        #8 COURIER
        $this->result['Couriers'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],               ['COURIER']),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],   ['COURIER']),
            'SG/TL'      => $this->sumCells(['I'],                    ['COURIER'])
                          + $this->sumCells(['I'],                     'STUD BY STUD SEPT2009'),
            'APP'        => $this->sumCells(['U', 'AC'],              ['COURIER'])
                          + $this->sumCells(['U', 'AC'],               'STUD BY STUD SEPT2009')
        ];

        #9 CONSULTANTS
        $this->result['Consultants'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],               ['CONSULTANTS']),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],   ['CONSULTANTS']),
            'SG/TL'      => $this->sumCells(['I'],                    ['CONSULTANTS']),
            'APP'        => $this->sumCells(['U', 'AC'],              ['CONSULTANTS'])
        ];

        #10 STUD BY STUD
        $this->result['Student-by-Student'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],               ['STUD BY STUD']),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],   ['STUD BY STUD']),
            'SG/TL'      => $this->sumCells(['I'],                    ['STUD BY STUD'],                 ['STUD BY STUD SEP2012']), // И чем им этот год не понравился?))
            'APP'        => $this->sumCells(['U', 'AC'],              ['STUD BY STUD'])
        ];

        #11 BIRTHDAY
        $this->result['Birthday action'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],               ['BIRTHDAY'],                     ['BIRTHDAY     PHONE DEC 2009', 'BIRTHDAY NON E-MAIL DEC 2012', 'BIRTHDAY NON E-MAIL DEC 2013', 'BIRTHDAY NON E-MAIL DEC 2014', 'BIRTHDAY NON E-MAIL DEC 2015', 'BIRTHDAY NON E-MAIL JAN 2016', 'BIRTHDAY NON E-MAIL JUL 2011', 'BIRTHDAY NON E-MAIL JUL 2014']),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'],   ['BIRTHDAY'],                     ['BIRTHDAY     PHONE DEC 2009', 'BIRTHDAY NON E-MAIL DEC 2012', 'BIRTHDAY NON E-MAIL DEC 2013', 'BIRTHDAY NON E-MAIL DEC 2014', 'BIRTHDAY NON E-MAIL DEC 2015', 'BIRTHDAY NON E-MAIL JAN 2016', 'BIRTHDAY NON E-MAIL JUL 2011', 'BIRTHDAY NON E-MAIL JUL 2014']),
            'SG/TL'      => $this->sumCells(['I'],                    ['BIRTHDAY'])
                          + $this->sumCells(['I'],                     'INTERNET     LET CENI2011 BD'),
            'APP'        => $this->sumCells(['U', 'AC'],              ['BIRTHDAY'])
        ];

        #12
        $this->result['INTERNET     SEARCH.MAIL.RU'] = [
            'prev_SG/TL' => $this->sumCells(['J'],        'INTERNET     SEARCH.MAIL.RU'),
            'prev_APP'   => $this->sumCells(['V', 'AD'],  'INTERNET     SEARCH.MAIL.RU'),
            'SG/TL'      => $this->sumCells(['I'],        'INTERNET     SEARCH.MAIL.RU'),
            'APP'        => $this->sumCells(['U', 'AC'],  'INTERNET     SEARCH.MAIL.RU')
        ];

        #13 - ...
        foreach ($this->singleResultRows as $rowKey){
            $this->result[$rowKey] = [
                'prev_SG/TL' => '',
                'prev_APP'   => '',
                'SG/TL'      => $this->sumCells(['I'],       $rowKey),
                'APP'        => $this->sumCells(['U', 'AC'], $rowKey)
            ];
        }


/*
 * Хуй у меня получилось сэкономить время, содрав формулы автоматом
 * Если слишколм долго вглядываться в этот отчёт, можно начать писать самому себе в комментах

        // Проходим по всем яйчейкам с православно спизженными из куска xlsx шаблона в виде xml формулами
        // Те из них, которые можно содрать автоматом, вычисляем
        for ($i = 13; $i < count(self::$RESULT_CAPTIONS); $i++){ // $RESULT_CAPTIONS нумерация с 0
            $row = $i + 8;
            $this->result[self::$RESULT_CAPTIONS[$i]] = [
                'prev_SG/TL' => isset($this->formulas[$row]['B']) ? $this->sumFormula($this->formulas[$row]['B']) : '',
                'prev_APP'   => isset($this->formulas[$row]['C']) ? $this->sumFormula($this->formulas[$row]['C']) : '',
                'SG/TL'      => isset($this->formulas[$row]['E']) ? $this->sumFormula($this->formulas[$row]['E']) : '',
                'APP'        => isset($this->formulas[$row]['F']) ? $this->sumFormula($this->formulas[$row]['F']) : '',
            ];
        }

*/

        # Link Exchange
        $this->result['Link Exchange'] = [
            'prev_SG/TL' => $this->sumCells(['J'],        'INTERNET     Link Exchange'),
            'prev_APP'   => $this->sumCells(['V', 'AD'],  'INTERNET     Link Exchange'),
            'SG/TL'      => $this->sumCells(['I'],        'INTERNET     Link Exchange'),
            'APP'        => $this->sumCells(['U', 'AC'],  'INTERNET     Link Exchange')
        ];


        // Этот ряд будет вычислен попозжа, пока вставляем его на своё место
        $this->result['Other channels'] = [
            'prev_SG/TL' => 0,
            'prev_APP'   => 0,
            'SG/TL'      => 0,
            'APP'        => 0
        ];


        # Unidentified others by phone
        $this->result['Unidentified others by phone'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],             'INFO-LINE'),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'], 'INFO-LINE'),
            'SG/TL'      => $this->sumCells(['I'],                  'INFO-LINE'),
            'APP'        => $this->sumCells(['U', 'AC'],            'INFO-LINE'),
        ];

        # Unidentified others written
        $this->result['Unidentified others written'] = [
            'prev_SG/TL' => $this->sumCells(['I', 'J'],             'LETTERS'),
            'prev_APP'   => $this->sumCells(['U', 'V', 'AC', 'AD'], 'LETTERS'),
            'SG/TL'      => $this->sumCells(['I'],                  'LETTERS'),
            'APP'        => $this->sumCells(['U', 'AC'],            'LETTERS'),
        ];


        # Other channels
        // Зависит от Unidentified others by phone и Unidentified others written
        $this->result['Other channels']['prev_SG/TL'] =
            $this->cell('Total', 'I')
          + $this->cell('Total', 'J')
          - $this->result['Print media adverts']            ['prev_SG/TL']
          - $this->result['Print media inserts']            ['prev_SG/TL']
          - $this->result['Remails']                        ['prev_SG/TL']
          - $this->result['Follow up']                      ['prev_SG/TL']
          - $this->result['Continuation graduates']         ['prev_SG/TL']
          - $this->result['Internet']                       ['prev_SG/TL']
          - $this->result['Internet Load']                  ['prev_SG/TL']
          - $this->result['Telemarketing']                  ['prev_SG/TL']
          - $this->result['Couriers']                       ['prev_SG/TL']
          - $this->result['Consultants']                    ['prev_SG/TL']
          - $this->result['Student-by-Student']             ['prev_SG/TL']
          - $this->result['Birthday action']                ['prev_SG/TL']
          - $this->result['Unidentified others by phone']   ['prev_SG/TL']
          - $this->result['Unidentified others written']    ['prev_SG/TL'];

        $this->result['Other channels']['prev_APP'] =
            $this->cell('Total', 'U')
          + $this->cell('Total', 'V')
          + $this->cell('Total', 'AC')
          + $this->cell('Total', 'AD')
          - $this->result['Print media adverts']            ['prev_APP']
          - $this->result['Print media inserts']            ['prev_APP']
          - $this->result['Remails']                        ['prev_APP']
          - $this->result['Follow up']                      ['prev_APP']
          - $this->result['Continuation graduates']         ['prev_APP']
          - $this->result['Internet']                       ['prev_APP']
          - $this->result['Internet Load']                  ['prev_APP']
          - $this->result['Telemarketing']                  ['prev_APP']
          - $this->result['Couriers']                       ['prev_APP']
          - $this->result['Consultants']                    ['prev_APP']
          - $this->result['Student-by-Student']             ['prev_APP']
          - $this->result['Birthday action']                ['prev_APP']
          - $this->result['Unidentified others by phone']   ['prev_APP']
          - $this->result['Unidentified others written']    ['prev_APP'];

        // Должно вычисляться в самую последнюю очередь, т.к. зависит от кучи предыдущих ячеек
        $this->result['Other channels']['SG/TL'] =
            $this->cell('Total', 'I')
          + self::$CELL_I1
          - array_sum(array_column($this->result, 'SG/TL'));

        $this->result['Other channels']['APP'] =
            $this->cell('Total', 'U')
          + $this->cell('Total', 'AC')
          - array_sum(array_column($this->result, 'APP'));



        // Да ну нахуй, нереально. Этот метод никогда не будет дописан до конца... Не верю в это
        return $this->result;
    }




    /**
     * Суммирование выбранных столбцов в диапазоне строк, кроме строк, исключённых из этого диапазона
     * @param array $cols Массив индексов столбцов, которые суммируются
     * @param string|array $range Диапазон строк, или массив таких диапазонов
     * @param array $keysExcept Массив исключённых из суммирования айдишников
     * @return int
     * @throws BaseException
     */
    public function sumCells($cols, $range, $keysExcept = [])
    {
        $sum = 0;

        // Если передана одна строка, значит это одинокий айдишник
        if (is_string($range)) {
            return $this->sumRow($range, $cols);
        }

        if (!is_array($range) || count($range) == 0) {
            throw new BaseException("Неправильный формат диапазона: " . var_export($range, true));
        }

        // Если передан массив из одного элемента, это имя группы айдишников, получаем её первый и последний ключ
        if (is_array($range) && count($range) == 1){
            if (!isset($this->rowGroups[$range[0]])){
                throw new BaseException("Неизвестная группа ключей {$range[0]}");
            }
            return $this->sumCells(
                $cols,
                [
                    reset($this->rowGroups[$range[0]]),
                    end($this->rowGroups[$range[0]])
                ],
                $keysExcept
            );
        }

        // Если передан массив диапазонов, рекурсивно собираем результат
        if (is_array($range[0])) {
            foreach ($range as $r) {
                $sum += self::sumCells($r, $cols, $keysExcept);
            }
            return $sum;
        }

        // Дальше ожидается только массив из двух элементов
        if (count($range) !== 2) {
            throw new BaseException("Неправильный формат диапазона: " . var_export($range, true));
        }

        // Чтобы диапазон можно было задавать и числами, и строками, изъебнёмся
        $startIndex = is_numeric($range[0]) ? intval($range[0]) : $this->keyIndexes[$range[0]];
        $endIndex   = is_numeric($range[1]) ? intval($range[1]) : $this->keyIndexes[$range[1]];

        // Проходим по диапазону
        for ($keyIndex = $startIndex; $keyIndex <= $endIndex; $keyIndex++) {
            // Если ключ в исключениях, не поверите - не суммируем
            if (in_array($this->keyIndexes[$keyIndex], $keysExcept)) {
                continue;
            }

            $sum += $this->sumRow($keyIndex, $cols);
        }

        return $sum;
    }




    /**
     * Суммирование выбранных ячеек одного ряда
     * @param string $rowKey Ключ ряда
     * @param array|int $cols Массив выбранных столбцов, или один столбец
     * @return int
     * @throws BaseException
     */
    public function sumRow($rowKey, $cols)
    {
        if (!is_array($cols)){
            return $this->cell($rowKey, $cols);
        }
        $sum = 0;
        // Проходим по всем заданным столбцам и суммируем ячейки
        foreach ($cols as $col) {
            $sum += $this->cell($rowKey, $col);
        }
        return $sum;
    }




    /**
     * Суммирование ячеек по формуле
     * @param string $formula Пожатый вид формулы - [столбец.ряд:столбец.ряд + ...]
     * @return string
     * @throws BaseException
     */
    public function sumFormula($formula){
        if (!$formula){
            return 0;
        }

        $sum = 0;

        // Разбиваем строку на слагаемые
        $elements = explode('+', $formula);
        foreach ($elements as $element){

            // Если в слагаемом задана отдельная ячейка, добавляем её к результату              $matches = [$element, столбец, ряд]
            if (preg_match('/^(\d{1,5})\\' . self::COL_ROW_DELIMITER . '(\d{1,5})$/', $element, $matches)) {
                $sum += $this->cell($matches[2], $matches[1]);

            // Если в слагаемом задан диапазон ячеек, получаем его данные и запускаем суммирование по диапазону
            } else if (preg_match('/^(\d{1,5})\\' . self::COL_ROW_DELIMITER . '(\d{1,5})\:(\d{1,5})\\' . self::COL_ROW_DELIMITER . '(\d{1,5})$/', $element, $matches)) {

                // Задаём координаты области и поехали                                 $matches = [$element, столбец, ряд, столбец, ряд]
                $sum += $this->sumCells(range($matches[1], $matches[3]), [$matches[2], $matches[4]]);


            // Особые подход для особых ячеек
            } else {

                //
                switch ($element){
                    // Ячейка I1 забита константой, которой Энштейну не хватило для доказательства Теории относительноси
                    case '0.-3':
                        $sum += self::$CELL_I1;
                        break;

                    // По нашей босяцкой жизни непременно что-то пойдёт не так
                    default:
                        throw new BaseException("Что-то не так со слагаемым $element... Нет, ты правда верил в то, что это сработает?");
                }
            }

        }

        return $sum;
    }





    /**
     * Вычисление суммы ячеек для выбранных строк и столбцов
     * @param array $colsIndexes Массив индексов столбцов, которые суммируются
     * @param array $rows Массив айдишников, по которым идёт суммирование
     * @return int
     * @throws BaseException
     */
    public function sumSelectCells($colsIndexes, $rows)
    {
        $sum = 0;

        // Проходим по всем строкам группы, проверяя, что они не в исключениях
        foreach ($rows as $rowKey){
            // Проходим по всем выбранным столбцам
            foreach ($colsIndexes as $col){
                $sum += $this->cell($rowKey, $col);
            }
        }

        return $sum;
    }




/** - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - Mr Hankey's christmas classics - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */



    /**
     * Выковыривание текста формул из файла шаблона
     * @param string $filename
     * @return bool
     */
    public function loadFormulas($filename){
        $xml = $xml = simplexml_load_file($filename);

        $rowCounter = 1;
        $this->formulas = [];
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
                        //$result[$rowCounter][$ch] =	isset($c->f) ? (string)$c->f : '';
                        if (isset($c->f) && (string)$c->f !== '') {
                            $this->formulas[$rowCounter][$ch{0}] = (string)$c->f;
                        }
                    }
                }
            }
            $rowCounter++;
            if ($rowCounter > 375) {
                // XML... ну вы понимаете...
                break;
            }
        }

        //var_dump($this->formulas);

        // Трахаем формулы страпоном
        $this->refactorFormulas();

        return true;
    }






    /**
     * Преобразование массива формул в формат (числовой индекс строки - числовой индекс столбца) в матрице рабочих данных
     * @return bool
     * @throws BaseException
     */
    protected function refactorFormulas()
    {

        // Проходим по всем рядам и столбцам
        foreach ($this->formulas as $rowIndex => $row) {
            foreach ($row as $colIndex => $cell) {
                if ($cell === '') {
                    continue;
                }
                // Убираем лишние символы, заменяем запятую (перечисление) на +
                $cell = str_replace('$', '', $cell);
                $cell = str_replace(')', '', $cell);
                $cell = str_replace('SUM(', '', $cell);
                $cell = str_replace(',', '+', $cell);

                // Заменяем символьное обозначение столбца на числовое
                preg_match_all('/([A-Z]{1,2})(\d{1,5})/i', $cell, $matches);
                foreach ($matches[0] as $i => $cl){
                    $col = $matches[1][$i];
                    $row = $matches[2][$i];
                    if (self::hasCol($col)){
                        //throw new BaseException("Столбец '$col' нам неизвестен");
                        continue;
                    }

                    $cell = str_replace(
                        $matches[0][$i],
                        self::colIndex($col) . self::COL_ROW_DELIMITER . ($row - self::ROW_INDEX_DIFF),
                        $cell
                    );
                }

                // В конце концов, заменяем исходную ячейку обработанной
                $this->formulas[$rowIndex][$colIndex] = $cell;
            }
        }
        return true;
    }






    /** Число строк в рабочем наборе */
    public function rowsCount()
    {
        return count($this->_cells);
    }



    /** Вывод в удобном виде формул */
    public function showFormulas()
    {
        $result = "[\n";
        foreach ($this->formulas as $k => $subArr) {
            $result .= "\t" . Filter::strPad("'$k'", 6) . " => [";
            $row = '';
            foreach ($subArr as $key => $cell) {
                $row .= $row ? ', ' : '';
                $row .= "'$key' => '$cell'";
            }
            $result .= "$row],\n";
        }
        return "$result\n]";
    }



    /** Вывод в удобном виде результата */
    public function showResult(){
        $result = '';
        reset(self::$RESULT_CAPTIONS);
        foreach ($this->result as $k => $subArr) {
            $result .= "\t" . Filter::strPad("'$k'", 32) . " => [ ";
            $row = '';
            foreach ($subArr as $key => $cell) {
                $row .= Filter::strPad("$cell", $key !== 'APP' ? 12 : 3);
            }
            $result .= "$row],\n";
        }
        return "[\n$result]";
    }


    /**
     * Вывод в удобном виде матрицы рабочих данных
     * @param bool $showEmpty
     * @return string
     */
    function showCells($showEmpty = true){
        $result = '';
        foreach ($this->_cells as $k => $v){
            if ($showEmpty || array_sum($v) !== 0) {
                $result .= Filter::strPad("'$k'", 32) . " => [ ";
                $row = '';
                foreach ($v as $i => $num) {
                    $num .= $i === count($v) - 1 ? '' : ', ';
                    $row .= Filter::strPad("$num", $i === count($v) - 1 ? 2 : 6);
                }
                $result .= "$row]";
                if ($k !== 'Total'){
                    $result .= "\n";
                }
            }
        }
        return $result;
    }




    /**
     * Числовой индекс столбца
     * @param string|array $cols
     * @return int|array
     * @throws BaseException
     */
    public static function colIndex($cols){
        if (is_array($cols)){
            return array_map([__CLASS__, 'colIndex'], $cols);
        }
        $cols = strtoupper($cols);
        if (!isset(self::$COL_INDEXES[$cols])){
            throw new BaseException("Неизвестный столбец '$cols'");
        }
        return self::$COL_INDEXES[$cols];
    }



    /**
     * Проверка существования столбца
     * @param string $col
     * @return bool
     */
    public static function hasCol($col){
        return isset(self::$COL_INDEXES[$col]);
    }
}