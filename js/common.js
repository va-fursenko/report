/**
 * Yippee-ki-yay, motherfucker!
 * It's a common js
 */




/**
 * Запись одной строки в лог
 * @param message Не поверите, та самая строка
 */
function logLine(message){
    if (message !== '') {
        var d = new Date();
        var t = (d.getHours() > 9 ? d.getHours() : '0' + d.getHours()) + ':' +
            (d.getMinutes() > 9 ? d.getMinutes() : '0' + d.getMinutes()) + ':' +
            (d.getSeconds() > 9 ? d.getSeconds() : '0' + d.getSeconds());
        $("#logPre").text($("#logPre").text() + '[' + t + '] ' + message + "\n");
    }else{
        $("#logPre").text($("#logPre").text() + "\n");
    }
}




/**
 * Завершение вычислений
 */
function onFinish(){
    $("#beginBtn").show();
    $("#logLoader").hide();
}



/**
 * Вывод в удобоваримой форме пользовательских данных *
 */
function showUserData(data)
{
    var blueRows = [
        "TOTAL THIS WEEK FROM WINTER CAMPAIGN",
        "TOTAL THIS WEEK FROM ALL OTHER CAMPAIGNS",
        "TOTAL THIS WEEK"
    ];
    var yellowRows = [
        "Internet",
        "Internet Load"
    ];
    var row, cell, table, tr, td;
    table = $('table#userDataTable');
    table.children('tbody').children('tr').remove();
    for (row in data) {
        tr = $('<tr>');

        if (row.match(/^internet.*/i)) {
            tr.addClass('bg-yellow');
        } else if (blueRows.indexOf(row) != -1) {
            tr.addClass('bg-blue');
        } else if (row == 'Other channels') {
            tr.addClass('b');
        }

        tr.append($('<td>').addClass('col-left').html(row));
        for (cell in data[row]) {
            td = $('<td>').addClass('res').html(data[row][cell]);
            if ((row == 'Other channels') && ((cell == 'SG/TL') || (cell == 'APP'))) {
                td.addClass('bg-blue');
            }
            tr.append(td);
        }
        table.append(tr);
    }
    $('#logPre').slideUp(500, function(){
        $('#userDataTable').slideDown(500);
    });

}


/**
 * Выполнение одного этапа задачи с логгированием результатов
 * @param act Действие для передачи в контроллер
 */
function nextStep(act){
    $.ajax({
        type: 'GET',
        url: '/controller.php?action=' + act,
        dataType: 'json',
        timeout: 240000,
        success: function (response) {
            if (response.success) {
                logLine(response.message == ''
                    ? "Завершено"
                    : "Завершено. " + response.message
                );

                // Если есть следующий этап, выполняем его
                if (typeof response.nextStep !== 'undefined'){
                    logLine('');
                    logLine(response.nextMessage);
                    nextStep(response.nextStep);

                // Прячем лоадер, показываем кнопку перезапуска и выводим удобный результат
                }else{
                    onFinish();
                    if (typeof response.userData !== 'undefined'){
                        showUserData(response.userData);
                    }
                }

            }else{
                logLine("# Произошла ошибка: " + response.message);
                onFinish();
            }
        },
        error: function() {
            logLine("# Произошла ошибка");
            onFinish();
        }
    });
}




/**
 * Действия при загрузке страницы
 */
$(window).load(function(){
    $(".notice-row").slideDown(500);
    $('#userDataTable').hide();

    // Старт отчёта
    $("#beginBtn").click(function(){
        $(".notice-row").slideUp();
        $('#userDataTable').slideUp();
        $("#beginBtn").hide();
        $("#logLoader").show();
        $("#logPre").text('');
        $('#logPre').show();

        logLine('# Импорт широкой матрицы');
        //logLine("# Вычисление отчёта");

        // Выполняем первый шаг
        nextStep('openFirst');
        //nextStep('process');
    })
});