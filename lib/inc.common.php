<?php
/**
 * Файл подключения общих модулей. Подразумевается, что конфиг подключён ранее
 * User: Виктор
 * Date: 01.03.2016
 * Time: 21:33
 */

/* Base libs */
require_once(CONFIG::ROOT . '/lib/class.BaseException.php');
require_once(CONFIG::ROOT . '/lib/class.Log.php');
require_once(CONFIG::ROOT . '/lib/class.Db.php');
require_once(CONFIG::ROOT . '/lib/class.Tpl.php');
require_once(CONFIG::ROOT . '/lib/class.ErrorHandler.php');




/* External libs */

/* PHPExcel */
require_once(CONFIG::ROOT . '/lib/external/PHPExcel/PHPExcel.php');
require_once(CONFIG::ROOT . '/lib/external/PHPExcel/PHPExcel/IOFactory.php');