<?php

require_once __DIR__ . DIRECTORY_SEPARATOR . 'result.captions.php';

/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 15.03.16
 * Time: 14:55
 */
class Report
{

    /** @const Константная ячейка */
    const CELL_I1 = 33777;

    /**
     * @const Разделитель столбца и ряда в обозначении ячейки в обработанной формуле
     * ОДИН СИМВОЛ
     */
    const COL_ROW_DELIMITER = '.'; // Точка пока смотрится просто удобнее в дампе



    /** @const Разница в индексе ряда между xls и матрицей рабочих данных */
    const ROW_INDEX_DIFF = 4; // Ячейка U8 из xls в нашей матрице имеет индекс [12, 8 - ROW_INDEX_DIFF]. Т.е. сохраняя её формулу, уменьшим индекс ряда на ROW_INDEX_DIFF



    /** @const Постоянные строки в первом файле. Расположены в конце и не участвуют в сравнении */
    const PERMANENT_ROWS = [
        'Телеобзвон',
        'REMAIL 1,2,3'
    ];



    // Инициализуется в result.captions.php Да, говнокод
    public static $RESULT_CAPTIONS = [];



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



    /** @const Связь между буквенными индексами */
    protected static $COL_INDEXES = ['I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH'];
    protected $COLS_INDEXES = [];



    /** @property array Рабочие данные */
    protected $_cells = [];



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
        // Вы же не собираетесь создавать два объекта класса подряд? :)
        $this->COLS_INDEXES = array_flip(self::$COL_INDEXES);

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

            // Соединяем текущий ряд из первой и второй матрицы
            $this->_cells[$keySecond] = array_merge($firstM[$keyFirst], $secondM[$keyFirst]);
            $this->keyIndexes[$rowIndex - 1] = $keySecond;
            $this->keyIndexes[$keySecond] = $rowIndex - 1;
        }

        // Дописываем фиксированные строки в конце первой матрицы нулями
        foreach (self::PERMANENT_ROWS as $index => $row){
            $this->_cells[$row] = array_pad($firstM[$row], MATR_FIRST_COLS + MATR_SECOND_COLS, 0);
            $this->keyIndexes[$rowIndex + $index] = $row;
            $this->keyIndexes[$row] = $rowIndex + $index;
        }

        // Человеку с фамилией Итого всегда достаётся больше других
        $this->_cells['Total'] = array_pad([], MATR_FIRST_COLS + MATR_SECOND_COLS, 0);
        foreach ($this->_cells['Total'] as $col => $value){
            $this->_cells['Total'][$col] = array_sum(array_column($this->_cells, $col));
        }

    }





    /**
     * Поэтапное почти что ручное вычисление отчёта
     * @return array Результат отчёта - $this->result
     */
    public function process(){
        // Сегодня 15 марта 2016 года, 17:00 мск. Меня зовут Виктор и я начал вручную собирать примерно 384 ряда по 4 ячейки
        $this->result = [];
        reset(self::$RESULT_CAPTIONS);

        // Проходим по всем яйчейкам с правславными (которые можно содрать автоматом) формулами и смотрим, что там есть
        for ($row = 8; $row <= 378; $row++){
            $this->result[each(self::$RESULT_CAPTIONS)['value']] = [
                'prev_SG/TL' => isset($this->formulas[$row]['B']) ? $this->sumFormula($this->formulas[$row]['B']) : '',
                'prev_APP'   => isset($this->formulas[$row]['C']) ? $this->sumFormula($this->formulas[$row]['C']) : '',
                'SG/TL'      => isset($this->formulas[$row]['E']) ? $this->sumFormula($this->formulas[$row]['E']) : '',
                'APP'        => isset($this->formulas[$row]['F']) ? $this->sumFormula($this->formulas[$row]['F']) : '',
            ];
        }

        // Плохая новость в том, что остались ряды, которые нужно собрать вручную, ибо их формулы ссылаются на вышевычисленные
        // Хорошая в том, что их не 380, а всего 9



        # Other channels
        // Зависит от Unidentified others by phone и Unidentified others written
        $this->result['Other channels']['prev_SG/TL'] =
            $this->cell('Total', $this->colIndex('I'))
            + $this->cell('Total', $this->colIndex('J'))
            - (
                $this->result['Print media adverts']              ['prev_SG/TL']
                + $this->result['Print media inserts']            ['prev_SG/TL']
                + $this->result['Remails']                        ['prev_SG/TL']
                + $this->result['Follow up']                      ['prev_SG/TL']
                + $this->result['Continuation graduates']         ['prev_SG/TL']
                + $this->result['Internet']                       ['prev_SG/TL']
                + $this->result['Internet Load']                  ['prev_SG/TL']
                + $this->result['Telemarketing']                  ['prev_SG/TL']
                + $this->result['Couriers']                       ['prev_SG/TL']
                + $this->result['Consultants']                    ['prev_SG/TL']
                + $this->result['Student-by-Student']             ['prev_SG/TL']
                + $this->result['Birthday action']                ['prev_SG/TL']
                + $this->result['Unidentified others by phone']   ['prev_SG/TL']
                + $this->result['Unidentified others written']    ['prev_SG/TL']
            );
        $this->result['Other channels']['prev_APP'] =
            $this->cell('Total', $this->colIndex('U'))
            + $this->cell('Total', $this->colIndex('V'))
            + $this->cell('Телеобзвон', $this->colIndex('AC'))
            + $this->cell('Телеобзвон', $this->colIndex('AD'))
            - (
                $this->result['Print media adverts']              ['prev_APP']
                + $this->result['Print media inserts']            ['prev_APP']
                + $this->result['Remails']                        ['prev_APP']
                + $this->result['Follow up']                      ['prev_APP']
                + $this->result['Continuation graduates']         ['prev_APP']
                + $this->result['Internet']                       ['prev_APP']
                + $this->result['Internet Load']                  ['prev_APP']
                + $this->result['Telemarketing']                  ['prev_APP']
                + $this->result['Couriers']                       ['prev_APP']
                + $this->result['Consultants']                    ['prev_APP']
                + $this->result['Student-by-Student']             ['prev_APP']
                + $this->result['Birthday action']                ['prev_APP']
                + $this->result['Unidentified others by phone']   ['prev_APP']
                + $this->result['Unidentified others written']    ['prev_APP']
            );


        # Unidentified others by phone
        $this->result['Unidentified others by phone']['prev_SG/TL'] =
            $this->cell('INFO-LINE', $this->colIndex('I'))
            + $this->cell('INFO-LINE', $this->colIndex('J'));

        $this->result['Unidentified others by phone']['prev_APP'] =
            $this->cell('INFO-LINE', $this->colIndex('U'))
            + $this->cell('INFO-LINE', $this->colIndex('V'))
            + $this->cell('INFO-LINE', $this->colIndex('AC'))
            + $this->cell('INFO-LINE', $this->colIndex('AD'));

        $this->result['Unidentified others by phone']['SG/TL'] =
            $this->cell('INFO-LINE', $this->colIndex('I'));

        $this->result['Unidentified others by phone']['APP'] =
            $this->cell('INFO-LINE', $this->colIndex('U'))
            + $this->cell('INFO-LINE', $this->colIndex('AC'));


        # Unidentified others written
        $this->result['Unidentified others written']['prev_SG/TL'] =
            $this->cell('LETTERS', $this->colIndex('I'))
            + $this->cell('LETTERS', $this->colIndex('J'));

        $this->result['Unidentified others written']['prev_APP'] =
            $this->cell('LETTERS', $this->colIndex('U'))
            + $this->cell('LETTERS', $this->colIndex('V'))
            + $this->cell('LETTERS', $this->colIndex('AC'))
            + $this->cell('LETTERS', $this->colIndex('AD'));

        $this->result['Unidentified others written']['SG/TL'] =
            $this->cell('LETTERS', $this->colIndex('I'));

        $this->result['Unidentified others written']['APP'] =
            $this->cell('LETTERS', $this->colIndex('U'))
            + $this->cell('LETTERS', $this->colIndex('AC'));


        //echo array_sum(array_column($this->result, 'SG/TL'));;

        // Должно вычисляться в самую последнюю очередь
        $this->result['Other channels']['SG/TL'] =
            $this->cell('Total', $this->colIndex('I'))
            + self::CELL_I1
            - array_sum(array_column($this->result, 'SG/TL'));

        $this->result['Other channels']['APP'] =
            $this->cell('Total', $this->colIndex('U'))
            + $this->cell('Телеобзвон', $this->colIndex('AC'))
            - array_sum(array_column($this->result, 'APP'));

        // prev_SG/TL
        // prev_APP
        // SG/TL
        // APP


        // Да ну нахуй, нереально. Этот метод никогда не будет дописан до конца... Не верю в это
        return $this->result;
    }




    /**
     * Суммирование выбранных столбцов в диапазоне строк, кроме строк, исключённых из этого диапазона
     * @param array $cols Массив индексов столбцов, которые суммируются
     * @param array $range Диапазон строк, или массив таких диапазонов
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
     * @param array $cols Массив выбранных столбцов
     * @return int
     * @throws BaseException
     */
    public function sumRow($rowKey, $cols)
    {
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
                        $sum += self::CELL_I1;
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
                    if (!isset($this->COLS_INDEXES[$col])){
                        //throw new BaseException("Столбец '$col' нам неизвестен");
                        continue;
                    }
                    $cell = str_replace(
                        $matches[0][$i],
                        $this->colIndex($col) . self::COL_ROW_DELIMITER . ($row - self::ROW_INDEX_DIFF),
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
        $result = "[\n";
        reset(self::$RESULT_CAPTIONS);
        foreach ($this->result as $k => $subArr) {
            $result .= "\t" . Filter::strPad("'$k'", 32) . " => [ ";
            $row = '';
            foreach ($subArr as $key => $cell) {
                $row .= Filter::strPad("$cell", $key !== 'APP' ? 16 : 3);
            }
            $result .= "$row],\n";
        }
        return "$result\n]";
    }


    /**
     * Вывод в удобном виде матрицы рабочих данных
     * @param bool $showEmpty
     * @return string
     */
    function showCells($showEmpty = true){
        $result = '';
        foreach ($this->_cells as $k => $v){
            if ($v[0] !== 0 || $showEmpty) {
                $result .= Filter::strPad("'$k'", 32) . " => [ ";
                $row = '';
                foreach ($v as $i => $num) {
                    $num .= $i === count($v) - 1 ? '' : ', ';
                    $row .= Filter::strPad("$num", $i === count($v) - 1 ? 2 : 6);
                }
                $result .= "$row]\n";
            }
        }
        return "[\n$result]";
    }




    /**
     * Числовой индекс столбца
     * @param string $char
     * @return int
     * @throws BaseException
     */
    public function colIndex($char){
        $char = strtoupper($char);
        if (!isset($this->COLS_INDEXES[$char])){
            throw new BaseException("Неизвестный столбец '$char'");
        }
        return $this->COLS_INDEXES[$char];
    }

}