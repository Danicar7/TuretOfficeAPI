/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.onReady(info => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("btnAnalizar").onclick = AnalizarFunction;
        document.getElementById("btnSuma").onclick = sumaPotencial;
    }
});

export async function run() {
    return Word.run(async context => {
        /**
         * Insert your Word code here
         */

        // insert a paragraph at the end of the document.
        const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

        // change the paragraph color to blue.
        paragraph.font.color = "blue";

        await context.sync();
    });
}


function AnalizarFunction() {
    Word.run(function(context) {

            var textoOriginal = context.document.getSelection();
            textoOriginal.load("text");
            return context.sync()
                .then(function() {
                    var palabras = textoOriginal.text;
                    context.document.body.insertParagraph("Texto copiado: " + palabras, "End");
                    consultaServidor(palabras);
                })
                .then(context.sync);
        })
        .catch(function(error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
}

function sumatoria() {
    var server = "http://turetapi.azurewebsites.net/";
    var op_num = { 'sum': [3, 4] };

    function update_var() {
        var n1 = parseInt(document.getElementById("n1").value); //parseFloat($("#n1").val());
        var n2 = parseInt(document.getElementById("n2").value); //parseFloat($("#n2").val());
        op_num['sum'] = [n1, n2];
    }

    var appdir = '/sum';
    var send_msg = "<p>Sending numbers</p>";
    var received_msg = "<p>Result returned</p>";
    update_var();
    console.log(send_msg);
    $('#message').html(send_msg);
    $.ajax({
        type: "POST",
        url: server + appdir,
        data: JSON.stringify(op_num),
        dataType: 'json'
    }).done(function(data) {
        console.log(data);
        $('#n3').val(data['sum']);
        $('#message').html(received_msg + data['msg']);
    });
}



function consultaServidor(texto) {
    var server = "http://127.0.0.1:5000"; //http://turetprueba1.azurewebsites.net/
    //var server = "http://analisisturet.azurewebsites.net";
    var appdir = '/analizar';
    var send_msg = "<p>Enviando texto</p>";
    var received_msg = "<p>Resultado retornado</p>";
    var vTexto = {
        'textoA': ''
    };

    function update_var() {
        var n1 = ($("#txtPrincipal").val());
        vTexto['textoA'] = [texto];
        //alert(vTexto['textoA']);
    }
    //alert(send_msg);
    //console.log(send_msg);
    update_var();
    $('#resultado').html(send_msg);

    $.ajax({
        type: "POST",
        url: server + appdir,
        data: JSON.stringify(vTexto),
        contentType: "application/x-www-form-urlencoded;charset=utf-8",
        dataType: 'json'
    }).done(function(data) {
        console.log(data);
        $('#txtResultado').val(data['textoA']);
        $('#resultado').html(received_msg + data['msg']);
        //$('#cal_sof').val(data['cal_sof']);
        //$('#nota_sof').val(data['nota_sof']);
        //$('#cal_var').val(data['cal_var']);
        //$('#nota_var').val(data['nota_var']);
        //$('#cal_den').val(data['cal_den']);
        //$('#nota_den').val(data['nota_den']);
        $('#cal_sof').html(data['cal_sof']);
        $('#nota_sof').html(data['nota_sof']);
        $('#cal_var').html(data['cal_var']);
        $('#nota_var').html(data['nota_var']);
        $('#cal_den').html(data['cal_den']);
        $('#nota_den').html(data['nota_den']);
    });
}

function sumaPotencial() {
    var server = "http://127.0.0.1:5000";
    var op_num = {
        'sum': [3, 4]
    };

    function update_var() {
        var n1 = parseFloat($("#n1").val());
        var n2 = parseFloat($("#n2").val());
        op_num['sum'] = [n1, n2];
    }
    $(function() {
        $("#sum").click(function() {
            var appdir = '/sum';
            var send_msg = "<p>Sending numbers</p>";
            var received_msg = "<p>Result returned</p>";
            update_var();
            console.log(send_msg);
            $('#message').html(send_msg);
            $.ajax({
                type: "POST",
                url: server + appdir,
                data: JSON.stringify(op_num),
                dataType: 'json'
            }).done(function(data) {
                console.log(data);
                $('#n3').val(data['sum']);
                $('#message').html(received_msg + data['msg']);
            });
        });
    });
}