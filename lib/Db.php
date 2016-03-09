<?php

/**
 * Сигнелтон PDO - работа с бд.
 */
class Db
{
    /**
     * @var null
     */
    private static $_instance = null;
    /**
     * @var \PDO
     */
    private $_db;

    /**
     * Запрещаем клонировать сингелтон
     */
    private function __clone()
    {
    }

    /**
     * метод инициализации
     */
    public function __construct()
    {
        try {
            $this->_db = new PDO ('mysql:host=localhost;dbname=homework',
                'root',
                '');
        } catch (Exception $e) {
            throw new Exception('Failed: connect to BD; Start application failed!', EXEPTION_CORE);
        }

        $this->_db->query("SET NAMES 'utf8'"); //Выставляем кодировку.
    }

    /**
     * @static
     * @return null
     */
    static function Instance()
    {

        if (self::$_instance == null) {
            self::$_instance = new Db();
        }
        return self::$_instance;
    }

    /**
     * @return \PDO
     */
    function getDb()
    {
        return $this->_db;
    }
}
