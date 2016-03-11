<?php 
	require_once($_SERVER['DOCUMENT_ROOT'] . DIRECTORY_SEPARATOR . 'config.php');
	require_once(CONFIG::ROOT . DIRECTORY_SEPARATOR . 'lib' . DIRECTORY_SEPARATOR . 'class.Log.php');

    if (isset($_GET['clear'])){
        fclose(fopen(__DIR__ . DIRECTORY_SEPARATOR . 'error.log', 'w'));
        header('Location: http://' . $_SERVER['SERVER_NAME'] . '/log/');
    }
?>
<!DOCTYPE html>
<html>

    <head>
        <title><?= CONFIG::PAGE_TITLE ?>&nbsp;/log/</title>
        <meta charset="<?= CONFIG::PAGE_CHARSET ?>">
        <script src="<?= CONFIG::HOST ?>js/jquery.js"></script>
        <script src="<?= CONFIG::HOST ?>js/bootstrap.js"></script>
        <link rel="stylesheet" href="<?= CONFIG::HOST ?>css/bootstrap.css">
        <link rel="stylesheet" href="<?= CONFIG::HOST ?>css/log.css">
    </head>

    <body>

        <script type='text/javascript'>
            jQuery( document ).ready(function() {
                jQuery('#scrollup img').mouseover( function(){
                    jQuery( this ).animate({opacity: 0.65},100);
                }).mouseout( function(){
                    jQuery( this ).animate({opacity: 1},100);
                }).click( function(){
                    window.scroll(0 ,0);
                    return false;
                });

                jQuery(window).scroll(function(){
                    if ( jQuery(document).scrollTop() > 0 ) {
                        jQuery('#scrollup').fadeIn('fast');
                    } else {
                        jQuery('#scrollup').fadeOut('fast');
                    }
                });
            });

            /*
            window.onload = function() { // после загрузки страницы

                var scrollUp = document.getElementById('scrollup'); // найти элемент

                scrollUp.onmouseover = function() { // добавить прозрачность
                    scrollUp.style.opacity=0.3;
                    scrollUp.style.filter  = 'alpha(opacity=30)';
                };

                scrollUp.onmouseout = function() { //убрать прозрачность
                    scrollUp.style.opacity = 0.5;
                    scrollUp.style.filter  = 'alpha(opacity=50)';
                };

                scrollUp.onclick = function() { //обработка клика
                    window.scrollTo(0,0);
                };

                // show button
                window.onscroll = function () { // при скролле показывать и прятать блок
                    if ( window.pageYOffset > 0 ) {
                        scrollUp.style.display = 'block';
                    } else {
                        scrollUp.style.display = 'none';
                    }
                };
            };
            */
        </script>

        <div id="logDiv">
            <div class="tab-header">
                <ul class="nav nav-pills" role="tablist">
                    <li class="active">
                        <a href="#e0" aria-controls="e1" role="tab" data-toggle="tab">Ошибки PHP</a>
                    </li>
                </ul>

                <div class="btn-container">
                    <a class="btn btn-danger" href="/log/index.php?clear">Очистить</a>
                </div>
            </div>

            <div class="tab-content">
                <div role="tabpanel" class="tab-pane fade active in" id="e0">
                    <?= Log::showLogFile('error.log') ?>
                </div>
            </div>
        </div>
        <div id="scrollup"><img alt="Прокрутить вверх" src="/img/up.png"></div>
    </body>
</html>


