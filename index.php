<?php

    // Конфиг
    require_once(__DIR__ . '/config.php');
    require_once(__DIR__ . '/lib/inc.common.php');
    require_once(__DIR__ . '/work/formulas.php');
    require_once(__DIR__ . '/work/import.xls.php');
    require_once(__DIR__ . '/work/formulas.refactor.php');

    // Рисуем страницу
    require_once(CONFIG::ROOT . DIRECTORY_SEPARATOR . CONFIG::TPL_DIR . '/layout.main.php');
