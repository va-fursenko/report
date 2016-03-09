<?php
/**
 * Created by PhpStorm.
 * User: viktor
 * Date: 04.03.16
 * Time: 14:59
 */

function makeDaddyHappy($filename){


    // Открываем файл
    $xls = PHPExcel_IOFactory::load($filename);
    // Устанавливаем индекс активного листа
    $xls->setActiveSheetIndex(1);
    // Получаем активный лист
    $sheet = $xls->getActiveSheet();

    return $sheet->getTitle();
}