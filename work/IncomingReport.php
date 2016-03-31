<?php

/**
 * Класс генерации месячной incoming статистики
 *
 * Конструктор принимает константу $I1 из шаблона отчёта,
 * а так же читает два csv файла 'WEEKC_CH.TXT' и 'reclama3.txt' в папке self::UPLOAD_DIR = '/protected/runtime/upload/statistic/'
 *
 * Результат может быть возвращён
 * - ассоциативным массивом (например, в фронт где его можно вывести в таблицу с помощью функции showUserData() из incomingReport.js),
 * - в xls-файле, который сохраняется в self::REPORT_DIR = '/assets/reports/'
 * в формате 'Month Incoming Statistics %YYYY %Month.xls';
 *
 * User: viktor
 * Date: 15.03.16
 * Time: 14:55
 */
class IncomingReport
{
    /** @const Директории для загрузки входных файлов и сохранения результата */
    const UPLOAD_DIR = '/protected/runtime/upload/statistic/';
    const REPORT_DIR = '/assets/reports/';

    /** @const Строка, с которой начинается полезная информация в файле отчёта */
    const TPL_FIRST_ROW = 3;
    /** @const Ширина первой (большой) матрицы */
    const MATR_FIRST_COLS = 18;
    /** @const  Ширина второй (маленькой) матрицы */
    const MATR_SECOND_COLS = 6;


    /** @const Группы айдишников, на которые забиваются все строки */
    public static $GROUP_KEYS = [
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

    /** @const Постоянные строки в первом файле. Расположены в конце и не участвуют в сравнении */
    public static $PERMANENT_ROWS = [
        'Телеобзвон',
        'REMAIL 1,2,3'
    ];

    /** @const Связь между буквенной индексацией столбцов в xls и в матрице рабочих данных */
    public static $COL_INDEXES = []; // Инициализуются простеньким выражением после декларации класса





    /**
     * Матрицы входных данных
     */
    public $wMatr = []; // WEEKC_CH
    public $rMatr = []; // reclama3


    /** @property array $group Сгруппированные айдишники
     * [
     *     'INTERNET' => ['INTERNET', 'INTERNET 2', 'INTERNET 9', ... ],
     *     ...
     */
    public $rowGroups = [];

    /** @property array Рабочие данные */
    public $cells = [];

    /** @property array Жёлтые ряды, которые вычисляются по одинаково простой схеме. Прогоним их через цикл */
    public $singleResultRows = [];

    /** @property array Результат отчёта */
    public $result = [];

    /**
     * @property array $keyIndexes
     * Порядковые (с 0) номера айдишников [0 => '21 CENTURY', 1 => 'ASET SG',...]
     * и обратные пары ['21 CENTURY' => 0, 'ASET SG' => 1,...]
     */
    public $keyIndexes = [];

    /** @const Константная ячейка */
    protected $cellI1 = -1000000; // Чтобы не забыть инициализовать





    /**
     * Конструктор класса - импорт широкой и узкой матрицы из указанных файлов, установка константы  "Скачанные с сайта"
     * @param int $cellI1 Значение ячейки I1 в шаблоне
     * @throws Exception
     */
    public function __construct(int $cellI1)
    {
        $WEEKC_CH_Path = Yii::getPathOfAlias('webroot') . self::UPLOAD_DIR . 'WEEKC_CH.TXT';
        $reclama3_Path = Yii::getPathOfAlias('webroot') . self::UPLOAD_DIR . 'reclama3.txt';
        $this->cellI1 = $cellI1;
        $this->wMatr = self::importMatrix($WEEKC_CH_Path, self::TPL_FIRST_ROW, self::MATR_FIRST_COLS);
        $this->rMatr = self::importMatrix($reclama3_Path, self::TPL_FIRST_ROW, self::MATR_SECOND_COLS);
    }



    /**
     * Универсальный геттер ячейки рабочих данных
     * @param int|string $row Числовой индекс или строковый ключ ряда
     * @param int $col Индекс столбца
     * @return int
     * @throws Exception
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
        if (!isset($this->cells[$row][$col])) {
            throw new Exception("Неверная адресация в массиве: [$row][$col]");
        }
        return $this->cells[$row][$col];
    }



    /**
     * Геттер матрицы рабочих данных
     */
    public function cells()
    {
        return $this->cells;
    }



    /**
     * Конструктор класса, принимает на вход две матрицы входных данных - 'широкую' и 'узкую' соответственно
     * @throws Exception
     */
    public function merge()
    {
        // За исключением двух строк в конце первого массива, они должны быть идентичными по ключам

        // Сравниваем длину с учётом постоянных элементов первой матрицы
        if (count($this->wMatr) - count(self::$PERMANENT_ROWS) !== count($this->rMatr)){
            throw new Exception('Размеры файлов не совпадают - ' . count($this->wMatr) . ' (в т.ч. ' . count(self::$PERMANENT_ROWS) . ' постоянные) и ' . count($this->rMatr));
        }

        // В первом массиве должны быть постоянные элементы
        foreach (self::$PERMANENT_ROWS as $row){
            if (!array_key_exists($row, $this->wMatr)){
                throw new Exception("Первый файл не имеет всех постоянных строк ['" . implode("', '", self::$PERMANENT_ROWS) . "']'");
            }
        }

        // Пройдём по второму массиву и сольём его со первым, проверяя, чтобы порядок ключей был одинаковым
        reset($this->wMatr);
        $rowIndex = 0;
        $this->keyIndexes = [];
        $this->cells = [];
        foreach ($this->rMatr as $keySecond => $value) {
            $rowIndex++;
            $keyFirst = each($this->wMatr)['key']; // Ключ первого массива на такой же позиции. Можно было его и засунуть внутрь ряда в отдельный столбец, но можно и так

            // Если на данной позиции ключи не равны, прерываем операцию и приунываем
            if ($keyFirst !== $keySecond) {
                throw new Exception("Файлы расходятся. Ряд $rowIndex, ключи ['$keyFirst'] и ['$keySecond']");
            }

            // Соединяем текущий ряд из первой и второй матрицы
            $this->cells[$keySecond] = array_merge($this->wMatr[$keyFirst], $this->rMatr[$keyFirst]);
            $this->keyIndexes[$rowIndex - 1] = $keySecond;
            $this->keyIndexes[$keySecond] = $rowIndex - 1;
        }

        // Дописываем фиксированные строки в конце первой матрицы нулями
        foreach (self::$PERMANENT_ROWS as $index => $row) {
            $this->cells[$row] = array_pad($this->wMatr[$row], self::MATR_FIRST_COLS + self::MATR_SECOND_COLS, 0);
            $this->keyIndexes[$rowIndex + $index] = $row;
            $this->keyIndexes[$row] = $rowIndex + $index;
        }

        // Человеку с фамилией Итого всегда достаётся больше других
        $this->cells['Total'] = array_pad([], self::MATR_FIRST_COLS + self::MATR_SECOND_COLS, 0);
        foreach ($this->cells['Total'] as $col => $value) {
            $this->cells['Total'][$col] = array_sum(array_column($this->cells, $col));
        }
        $this->keyIndexes['Total'] = count($this->keyIndexes) - 1;
        $this->keyIndexes[count($this->keyIndexes) - 1] = 'Total';

        // Проходим по матрице и раскладываем строки по группам айдишников
        $this->rowGroups = [];
        foreach ($this->cells as $key => $row) {
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

            // Исключаем всех, кроме исключений %) Ну не создавать же для одного айдишника массив с белым списком, или писать его в сравнение if(...)
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
        $this->singleResultRows = [];
        foreach ($this->cells as $key => $row) {
            if (!in_array($key, $exceptRows) && $this->sumCells(['I', 'U', 'AC'], $key) > 0) {
                $this->singleResultRows[] = $key;
            }
        }

    }



    /**
     * Импорт данных их json-строки
     * @param string $str
     */
    public function jsonImport($str)
    {
        $str = json_decode($str, true);
        $this->cells = $str['_cells'];
        $this->result = $str['result'];
        $this->keyIndexes = $str['keyIndexes'];
        $this->rowGroups = $str['rowGroups'];
        //$this->formulas = $str['formulas'];
        $this->singleResultRows = $str['singleResultRows'];
    }



    /**
     * Поэтапное почти что ручное вычисление отчёта
     * @return array Результат отчёта - $this->result
     */
    public function process()
    {
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
            'SG/TL'      => $this->sumCells(['I'],         'INTERNET     LOAD') + $this->cellI1,
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


        // Должно вычисляться в самую последнюю очередь, т.к. зависит от предыдущих ячеек
        $this->result['Other channels']['SG/TL'] =
            $this->cell('Total', 'I')
            + $this->cellI1
            - array_sum(array_column($this->result, 'SG/TL'));

        $this->result['Other channels']['APP'] =
            $this->cell('Total', 'U')
            + $this->cell('Total', 'AC')
            - array_sum(array_column($this->result, 'APP'));


        // Самое важное в отчёте, чтобы две какие-то ячейки были по нулям
        if ($this->result['Other channels']['SG/TL'] !== 0 || $this->result['Other channels']['APP'] !== 0) {
            return [
                'success'   => false,
                'message'   => "Отчёт не сошёлся! Other channels[SG/TL] = {$this->result['Other channels']['SG/TL']}, Other channels[APP] = {$this->result['Other channels']['APP']}"
            ];
        }


        // Вычисляем итоговые ряды результата

        # TOTAL THIS WEEK FROM WINTER CAMPAIGN                             WINTER IS COMING
        foreach (['prev_SG/TL', 'prev_APP'] as $col) {
            $this->result['TOTAL THIS WEEK FROM WINTER CAMPAIGN'][$col] =
                $this->result['Print media adverts']             [$col]
                + $this->result['Print media inserts']           [$col]
                + $this->result['Remails']                       [$col]
                + $this->result['Follow up']                     [$col]
                + $this->result['Continuation graduates']        [$col]
                + $this->result['Internet']                      [$col]
                + $this->result['Internet Load']                 [$col]
                + $this->result['Telemarketing']                 [$col]
                + $this->result['Couriers']                      [$col]
                + $this->result['Consultants']                   [$col]
                + $this->result['Student-by-Student']            [$col]
                + $this->result['Birthday action']               [$col]
                + $this->result['Unidentified others by phone']  [$col]
                + $this->result['Unidentified others written']   [$col];
        }
        foreach (['SG/TL', 'APP'] as $col) {
            $this->result['TOTAL THIS WEEK FROM WINTER CAMPAIGN'][$col] =
                array_sum(array_column($this->result, $col))
                - $this->result['Link Exchange']                 [$col]
                - $this->result['Other channels']                [$col];
        }

        # TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS
        $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS']   ['prev_SG/TL'] =
            $this->result['Other channels']                         ['prev_SG/TL']
            + $this->cell('Total', 'Q');
        $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS']   ['prev_APP'] =
            $this->result['Other channels']                         ['prev_APP']
            + $this->cell('Total', 'Y')
            + $this->cell('Total', 'AG');
        $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS']   ['SG/TL'] =
            $this->result['Other channels']                         ['SG/TL']
            + $this->cell('Total', 'Q')
            + $this->cell('Total', 'J');
        $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS']   ['APP'] =
            $this->result['Other channels']                         ['APP']
            + $this->cell('Total', 'Y')
            + $this->cell('Total', 'V')
            + $this->cell('Total', 'AD')
            + $this->cell('Total', 'AG');

        # TOTAL THIS WEEK
        $this->result['TOTAL THIS WEEK']                                ['prev_SG/TL'] =
            $this->result['TOTAL THIS WEEK FROM WINTER CAMPAIGN']       ['prev_SG/TL']
            + $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS']   ['prev_SG/TL']
            + $this->cellI1;

        $this->result['TOTAL THIS WEEK']                                ['prev_APP'] =
            $this->result['TOTAL THIS WEEK FROM WINTER CAMPAIGN']       ['prev_APP']
            + $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS']   ['prev_APP'];

        $this->result['TOTAL THIS WEEK']                                ['SG/TL'] =
            $this->result['TOTAL THIS WEEK FROM WINTER CAMPAIGN']       ['SG/TL']
            + $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS']   ['SG/TL'];

        $this->result['TOTAL THIS WEEK']                                ['APP'] =
            $this->result['TOTAL THIS WEEK FROM WINTER CAMPAIGN']       ['APP']
            + $this->result['TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS'] ['APP'];


        // Казалось, этот метод никогда не будет дописан до конца...
        return [
            'success'   => true,
            'message'   => 'Отчёт выполнен'
        ];
    }




    /**
     * Суммирование выбранных столбцов в диапазоне строк, кроме строк, исключённых из этого диапазона
     * @param array $cols Массив индексов столбцов, которые суммируются
     * @param string|array $range Диапазон строк, или массив таких диапазонов
     * @param array $keysExcept Массив исключённых из суммирования айдишников
     * @return int
     * @throws Exception
     */
    public function sumCells($cols, $range, $keysExcept = [])
    {
        $sum = 0;

        // Если передана одна строка, значит это одинокий айдишник
        if (is_string($range)) {
            return $this->sumRow($range, $cols);
        }

        if (!is_array($range) || count($range) == 0) {
            throw new Exception("Неправильный формат диапазона: " . var_export($range, true));
        }

        // Если передан массив из одного элемента, это имя группы айдишников, получаем её первый и последний ключ
        if (is_array($range) && count($range) == 1) {
            if (!isset($this->rowGroups[$range[0]])) {
                throw new Exception("Неизвестная группа ключей {$range[0]}");
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
            throw new Exception("Неправильный формат диапазона: " . var_export($range, true));
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
     * @throws Exception
     */
    public function sumRow($rowKey, $cols)
    {
        if (!is_array($cols)) {
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
     * Вычисление суммы ячеек для выбранных строк и столбцов
     * @param array $colsIndexes Массив индексов столбцов, которые суммируются
     * @param array $rows Массив айдишников, по которым идёт суммирование
     * @return int
     * @throws Exception
     */
    public function sumSelectCells($colsIndexes, $rows)
    {
        $sum = 0;
        // Проходим по всем строкам группы, проверяя, что они не в исключениях
        foreach ($rows as $rowKey) {
            // Проходим по всем выбранным столбцам
            foreach ($colsIndexes as $col) {
                $sum += $this->cell($rowKey, $col);
            }
        }
        return $sum;
    }



    /**
     * Выполнение отчёта и возврат массива пользовательских данных
     * @return array
     */
    public function exec(){
        $this->merge();
        $this->process();
        return $this->getResultUserData();
    }




    /** - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - Mr Hankey's christmas classics - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - */


    /** Число строк в рабочем наборе */
    public function rowsCount()
    {
        return count($this->cells);
    }



    /** Вывод в удобном текстовом виде результата */
    public function showResult()
    {
        $result = '';
        foreach ($this->result as $k => $subArr) {
            $result .= "\t" . str_pad("'$k'", 42) . " => [ ";
            $row = '';
            foreach ($subArr as $key => $cell) {
                $row .= str_pad("$cell", $key !== 'APP' ? 12 : 4);
            }
            $result .= "$row],\n";
        }
        return "[\n$result]";
    }


    /**
     * Вывод в удобном текстовом виде матрицы рабочих данных
     * @param bool $showEmpty
     * @return string
     */
    function showCells($showEmpty = true)
    {
        $result = '';
        foreach ($this->cells as $k => $v) {
            if ($showEmpty || array_sum($v) !== 0) {
                $result .= str_pad("'$k'", 32) . " => [ ";
                $row = '';
                foreach ($v as $i => $num) {
                    $num .= $i === count($v) - 1 ? '' : ', ';
                    $row .= str_pad("$num", $i === count($v) - 1 ? 2 : 6);
                }
                $result .= "$row]";
                if ($k !== 'Total') {
                    $result .= "\n";
                }
            }
        }
        return $result;
    }



    /**
     * Числовой индекс столбца
     * @param string|array $col
     * @return int|array
     * @throws Exception
     */
    public static function colIndex($col)
    {
        if (is_array($col)) {
            return array_map([__CLASS__, 'colIndex'], $col);
        }
        $col = strtoupper($col);
        if (!isset(self::$COL_INDEXES[$col])) {
            throw new Exception("Неизвестный столбец '$col'");
        }
        return self::$COL_INDEXES[$col];
    }



    /**
     * Проверка существования столбца
     * @param string $col
     * @return bool
     */
    public static function hasCol($col)
    {
        return isset(self::$COL_INDEXES[strtoupper($col)]);
    }




    /**
     * Выжимка из результата полезных для пользователя данных в виде ассоциативного массива
     * @return array
     */
    public function getResultUserData()
    {
        $result = [];
        $fixedRows = [
            'Print media adverts',
            'Print media inserts',
            'Remails',
            'Follow up',
            'Continuation graduates',
            'Internet',
            'Internet Load',
            'Telemarketing',
            'Couriers',
            'Consultants',
            'Student-by-Student',
            'Birthday action',
            'TOTAL THIS WEEK FROM WINTER CAMPAIGN',
            'TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS',
            'TOTAL THIS WEEK',
        ];
        foreach ($this->result as $key => $row) {
            if (in_array($key, $fixedRows) || array_sum($row) > 0) {
                $result[$key]['SG/TL'] = $row['SG/TL'];
                $result[$key]['APP'] = $row['APP'];
            }
        }
        return $result;
    }



    /**
     * Выгрузка результата в xls
     * @return string|false Полное имя сформированного файла, или false
     */
    public function getXLSXResult(){
        $filePath = self::REPORT_DIR . 'Month Incoming Statistics ' . date('Y') . ' ' . date('F') . '.xls';
        $xls = new PHPExcel();
        $xls->setActiveSheetIndex(0);
        $sheet = $xls->getActiveSheet();

        // Заполняем постоянные ячейки
        $sheet->setTitle('Месячная incoming-статистика')
            ->setCellValue('A1', 'Weekly incoming per channel')
            ->setCellValue('A2', 'ESCC:')
            ->setCellValue('B2', 'Russia')
            ->setCellValue('B3', 'month:')
            ->setCellValue('C3', strtolower(date('M')) . ' ' . date('Y'))
            ->setCellValue('B4', 'RESPONSES RECEIVED per CHANNEL')
            ->setCellValue('B5', "SG/TL\nthis month")
            ->setCellValue('C5', "Applications\nthis month")
            ->setCellValue('A6', "RESPONSES DIVIDED PER CHANNEL\nUSED IN CURRENT CAMPAIGN");

        $result = $this->getResultUserData();
        // Числовой счётчик ряда нужен для доступа к рядам excel, потому что $result ассоциативный
        $rowCounter = 7;

        foreach($result as $key => $row){

            $sheet->setCellValue('A' . $rowCounter, $key)
                  ->setCellValue('B' . $rowCounter, $row['SG/TL'])
                  ->setCellValue('C' . $rowCounter, $row['APP']);

            $sheet->getRowDimension($rowCounter)
                ->setRowHeight(20);

            if (stripos($key, 'INTERNET') === 0){
                $sheet->getStyle('A' . $rowCounter)
                    ->getFill()
                    ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
                    ->getStartColor()
                    ->setRGB('FFFF99');
            }

            $rowCounter++;
        }
        $rowCounter--;

        // Форматируем документ
        $sheet->getColumnDimension('A')
            ->setWidth(30);

        $sheet->getColumnDimension('B')
            ->setWidth(18);

        $sheet->getColumnDimension('C')
            ->setWidth(18);

        $sheet->getRowDimension(1)
            ->setRowHeight(30);

        $sheet->getRowDimension(2)
            ->setRowHeight(30);

        $sheet->getRowDimension(3)
            ->setRowHeight(30);

        $sheet->getRowDimension(5)
            ->setRowHeight(30);

        $sheet->getRowDimension(6)
            ->setRowHeight(30);

        // Объединяем некоторые ячейки
        $sheet->mergeCells('A1:B1');
        $sheet->mergeCells('B4:C4');
        $sheet->mergeCells('A6:C6');

        // Высота строк в результате
        $sheet->getRowDimension($rowCounter - 2)
            ->setRowHeight(30);

        $sheet->getRowDimension($rowCounter - 1)
            ->setRowHeight(30);

        // Окончательное оформление ячеек

        // Стиль по умолчанию
        $sheet->getStyle('A1:C' . $rowCounter)
            ->applyFromArray([
                'font'      => [
                    'name' => 'Times New Roman',
                    'size' => 12,
                ],
                'alignment' => [
                    'vertical' => PHPExcel_Style_Alignment::VERTICAL_CENTER,
                ],
                'borders' => [
                    'allborders' => [
                        'style' => PHPExcel_Style_Border::BORDER_THIN,
                        'color' => [
                            'rgb' => 'C0C0C0'
                        ],
                    ]
                ]
            ]);

        // Двойные границы ячеек
        $styleBorderDouble = [
            'style' => PHPExcel_Style_Border::BORDER_DOUBLE,
            'color' => [
                'rgb' => '696969'
            ]
        ];

        $styleAllBordersDouble = ['borders' => ['allborders' => $styleBorderDouble,]];
        $styleRightBorderDouble = ['borders' => ['right' => $styleBorderDouble,]];

        $sheet->getStyle('B4:C6')
            ->applyFromArray($styleAllBordersDouble);

        $sheet->getStyle('B4:C6')
            ->applyFromArray($styleAllBordersDouble);

        $sheet->getStyle('A6:C6')
            ->applyFromArray($styleAllBordersDouble);

        $sheet->getStyle('A7:A' . $rowCounter)
            ->applyFromArray($styleRightBorderDouble);

        $sheet->getStyle('C7:C' . $rowCounter)
            ->applyFromArray($styleRightBorderDouble);

        $sheet->getStyle('A' . ($rowCounter - 2) . ':C' . $rowCounter)
            ->applyFromArray($styleAllBordersDouble);

        // Синий фон рядов с результатом
        $sheet->getStyle('A' . ($rowCounter - 2) . ':C' . $rowCounter)
            ->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()
            ->setRGB('d0e3f7');

        // Серый фон контрольного ряда
        $sheet->getStyle('A' . ($rowCounter - 4) . ':C' . ($rowCounter - 4))
            ->getFill()
            ->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
            ->getStartColor()
            ->setRGB('e1e1e1');

        // Выравнивания и прочее
        $sheet->getStyle('A1:C6')
            ->getFont()
            ->setBold(true);

        $sheet->getStyle('A2')
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

        $sheet->getStyle('A1:C3')
            ->getFont()
            ->setSize(16);

        $sheet->getStyle('A6')
            ->getAlignment()
            ->setWrapText(true);

        $sheet->getStyle('B2')
            ->getFont()
            ->setItalic(true);

        $sheet->getStyle('B2')
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_LEFT);

        $sheet->getStyle('B3')
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

        $sheet->getStyle('B4')
            ->getAlignment()
            ->setWrapText(true);

        $sheet->getStyle('B4:C5')
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        $sheet->getStyle('B5:C5')
            ->getAlignment()
            ->setWrapText(true);

        $sheet->getStyle('A7:A' . ($rowCounter - 3))
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

        $sheet->getStyle('A' . ($rowCounter - 2) . ':A' . $rowCounter)
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        $sheet->getStyle('A' . ($rowCounter - 3) . ':A' . $rowCounter)
            ->getAlignment()
            ->setWrapText(true);

        $sheet->getStyle('B7:C' . $rowCounter)
            ->getAlignment()
            ->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

        $sheet->getStyle('A' . ($rowCounter - 2) . ':C' . $rowCounter)
            ->getFont()
            ->setBold(true);

        $sheet->getStyle('A' . ($rowCounter - 4) . ':C' . ($rowCounter - 4))
            ->getFont()
            ->setBold(true);

        // Выводим файл пользвоателю
        $objWriter = new PHPExcel_Writer_Excel5($xls);
        $objWriter->save(Yii::getPathOfAlias('webroot') . $filePath);
        return $filePath;
    }



    /**
     * Импорт матрицы из указанного файла
     * @param string $filePath  Имя файла
     * @param int $rowStartFrom Ряд, с которого начинается чтение
     * @param int $colCount     Ширина считываемой матрицы
     * @return mixed
     * @throws Exception
     */
    protected static function importMatrix($filePath, $rowStartFrom, $colCount)
    {
        if (!is_readable($filePath)) {
            throw new Exception("Файл отчёта $filePath недоступен");
        }
        $fileHandle = fopen($filePath, 'r');
        $rowCounter = 0;
        $matrix = [];
        $row = [true];
        $rowKey = '';

        // Пропускаем первые ряды
        while ($rowCounter < $rowStartFrom && isset($row[0])) {
            $row = fgetcsv($fileHandle, 1000, "\t");
            $rowCounter++;
            $row = array_map('trim', $row);
            $rowKey = iconv('cp866', 'utf-8', $row[0]);
        }

        // Проходим по всем остальным рядам
        while (count($row) > 1 && $rowKey != 'Total') {
            // Проверяем число столбцов в файле - первое поле с айдишником и полезные данные
            if (count($row) !== $colCount + 1){
                throw new Exception("Число столбцов файла '$filePath' не совпадает с ожидаемым: " . count($row) . ' вместо ' . ($colCount + 1));
            }

            // Все айдишники должны быть уникальными
            if (isset($matrix[$rowKey])){
                throw new Exception("Ключ '$rowKey' повторяется в строке $rowCounter");
            }

            // Выкидываем первый элемент, потому, что он используется в качестве ключа
            array_shift($row);
            $matrix[$rowKey] = $row;

            // Получаем следующий ряд
            $row = array_map('trim', fgetcsv($fileHandle, 1000, "\t", '"'));
            $rowKey = iconv('cp866', 'utf-8', $row[0]);
            $rowCounter++;
        }

        return $matrix;
    }
}




/**
 * Действия в модели - некошерная практика, но по сути это просто константный массив.
 * Просто не хочется руками забивать числовые индексы ячеек типа ['I'=>0, 'J'=>1,... 'AH'=>23], они ведь могут измениться
 */
IncomingReport::$COL_INDEXES = array_flip(['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH']);