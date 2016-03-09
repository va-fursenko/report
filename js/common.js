/**
 * Yippee-ki-yay, motherfucker!
 * It's a common js
 */




/**
 * Запись одной строки в лог
 * @param message Не поверите, та самая строка
 */
function logLine(message){
    var d = new Date();
    var t = (d.getHours() > 9 ? d.getHours() : '0' + d.getHours()) + ':' +
            (d.getMinutes() > 9 ? d.getMinutes() : '0' + d.getMinutes()) + ':' +
            (d.getSeconds() > 9 ? d.getSeconds() : '0' + d.getSeconds());
    $("#logPre").text($("#logPre").text() + '[' + t + '] ' + message + "\n");
}



/**
 * Выполнение одного этапа задачи с логгированием результатов
 * @param act Действие для передачи в контроллер
 */
function nextStep(act){
    $("#logLoader").show();
    $.ajax({
        type: 'GET',
        url: '/controller.php?action=' + act,
        dataType: 'json',
        success: function (data) {
            if (data.success) {
                logLine(data.message == ''
                    ? "Завершено"
                    : "Завершено: " + data.message
                );
                if (typeof data.nextStep !== 'undefined'){
                    logLine(data.nextMessage);
                    nextStep(data.nextStep);
                }
            }else{
                logLine("# Произошла ошибка: " + data.message);
            }
            $("#logLoader").slideUp();
        },
        error: function() {
            logLine("# Произошла ошибка");
            $("#logLoader").slideUp();
        }
    });
}




/**
 * Действия при загрузке страницы
 */
$(window).load(function(){
    logLine('# Импорт списка айдишников');

    // Выполняем первый шаг
    nextStep('openXLS');
});