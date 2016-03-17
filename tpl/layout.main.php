<!DOCTYPE html>
<html>

<head>
    <title><?= CONFIG::PAGE_TITLE ?></title>
    <meta charset="<?= CONFIG::PAGE_CHARSET ?>">
    <!-- js -->
    <script src="<?= CONFIG::HOST ?>/js/jquery.js"></script>
    <script src="<?= CONFIG::HOST ?>/js/bootstrap.js"></script>
    <script src="<?= CONFIG::HOST ?>/js/common.js"></script>
    <!-- css -->
    <link rel="stylesheet" href="<?= CONFIG::HOST ?>/css/bootstrap.css">
    <link rel="stylesheet" href="<?= CONFIG::HOST ?>/css/common.css">
</head>

<body>
    <div class="container" style="margin-top: 100px;">
        <div class="row log-row">
            <h4 class="log-caption">Отчёт по incoming статистике</h4>
            <img class="log-caption log-loader" id="logLoader" src="img/loader.gif">
            <a id="beginBtn" class="log-caption btn btn-primary" href="javascript:void(0);">Красная кнопка</a>
        </div>

<!--
        <div class="row notice-row">
            <span><sup class="text-danger">*</sup>Скрипт может выполняться значительное время. Возможно, придётся настроить web-сервер и PHP, а так же, дать права записи на папку data/ в проекте</span>
<pre class="notice-row">
NGNIX: http              { keepalive_timeout    240; }
       location ~ \.php$ { fastcgi_read_timeout 240; }

PHP: max_execution_time = 240
</pre>
        </div>
    </div>
-->
</body>
</html>
