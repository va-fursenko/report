<div class="row log-row">
    <h4 class="log-caption">Отчёт по чему-то важному</h4>
    <img class="log-caption log-loader" id="logLoader" src="img/loader.gif">
    <a id="beginBtn" class="log-caption btn btn-primary" href="javascript:void(0);">Создать отчёт</a>
</div>
<div class="row">
    <pre class="log-container" id="logPre">{lines}</pre>
</div>
<div class="row notice-row">
    <span><sup class="text-danger">*</sup>Скрипт может выполняться значительное время. Возможно, придётся настроить web-сервер и PHP, а так же дать права записи на папку data/ в проекте</span>

<pre class="notice-row">
NGNIX: http              { keepalive_timeout    240; }
       location ~ \.php$ { fastcgi_read_timeout 240; }

PHP: max_execution_time = 240
</pre>

</div>