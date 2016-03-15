<?php

/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 15.03.16
 * Time: 14:48
 */
class CellSum
{

    public $name = '';

    public $prevSGTL = '';
    public $prevAPP = '';
    public $SGTL = '';
    public $APP = '';



    /**
     * Конструктор класса. Создаём его из объекта матрицы ячеек и имени групп айдишников
     * @param Report $cells Объект матрицы
     * @param string $groupName Имя группы айдишников
     * @param array $colsIndexes Массив индексов столбцов, которые суммируются
     * @param array $exceptRows Массив айдишников исключений, которые не учитываются
     */
    public function __construct($cells, $groupName, $colsIndexes, $exceptRows)
    {

    }
}

