<?php

    // Конфиг
    require_once(__DIR__ . '/config.php');
    require_once(__DIR__ . '/lib/inc.common.php');

    require_once(__DIR__ . '/work/import.xls.php');
    require_once(__DIR__ . '/work/class.Report.php');

    // Рисуем страницу
    require_once(CONFIG::ROOT . DIRECTORY_SEPARATOR . CONFIG::TPL_DIR . '/layout.main.php');


$report = new Report(
    json_decode(file_get_contents(XLS_ROOT . XLS_FIRST), true),
    json_decode(file_get_contents(XLS_ROOT . XLS_SECOND), true)
);

//$report->loadFormulas(CONFIG::ROOT . '/data/sheet2.xml');

$report->process();
/*
$expr = '/^(\d{1,5})\\' . Report::COL_ROW_DELIMITER . '(\d{1,5})\:(\d{1,5})\\' . Report::COL_ROW_DELIMITER . '(\d{1,5})$/';
$el = '0.2146:1.2165';
if (preg_match($expr, $el, $matches)){
    var_dump($matches);
}
*/



// Грузим из файла формулы
?>
<pre class='log-container' style='margin-top:20px; margin-bottom:20px'>
<?= $report->showResult(); ?>
</pre>

<pre class="log-header">                                    [ I,    J,    K,    L,    M,    N,    O,    P,    Q,    R,    S,    T,    U,    V,    W,    X,    Y,    Z,   AC,   AD,   AE,   AF,   AG,   AH ]</pre>
<pre class="log-container" id="logPre">
<?php

    echo $report->showCells();

?>
</pre>
<pre class="log-header">                                    [ I,    J,    K,    L,    M,    N,    O,    P,    Q,    R,    S,    T,    U,    V,    W,    X,    Y,    Z,   AC,   AD,   AE,   AF,   AG,   AH ]</pre>

