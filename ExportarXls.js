window.onload = function () {
    
    var charset = 'utf-8';
    var nmGuia = 'NOME DA GUIA';
    var nmArquivo = 'NOME DO ARQUIVO';
    var uri = `data:application/vnd.ms-excel;charset=${charset};base64,`;
    var template = '<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">';
    template += `<meta http-equiv="content-type" content="application/vnd.ms-excel; charset=${charset}"><head><!--[if gte mso 9]><xml><x:ExcelWorkbook><x:ExcelWorksheets>`;
    template += '<x:ExcelWorksheet><x:Name>{worksheet}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets>';
    template += '</x:ExcelWorkbook></xml><![endif]--></head><body><table>{table}</table></body></html>';
    var base64 = function (s) {
        return window.btoa(unescape(encodeURIComponent(s)));
    };
    var format = function (s, c) {
        return s.replace(/{(\w+)}/g, function (m, p) {
            return c[p];
        });
    };

    // seleciona tabela
    var tabela = document.querySelector('table');

    //Nomeia guia e pega CONTEÚDO da tabela, tudo DENTRO das tags <table></table>
    var ctx = {
        worksheet: `${nmGuia}`,
        table: tabela.innerHTML
    };

    //objeto(<a>) para executar o download
    var link = document.createElement("a");

    //NECCESSARIO o append no BODY do DOCUMENT
    document.body.appendChild(link);

    // se o navegador for IE salva como blob, testado no IE10 e IE11
    var browser = window.navigator.appVersion;
    if ((browser.indexOf('Trident') !== -1 && browser.indexOf('rv:11') !== -1) ||
        (browser.indexOf('MSIE 10') !== -1)) {
        var builder = new window.MSBlobBuilder();
        builder.append(uri + format(template, ctx));
        var blob = builder.getBlob(`data:application/vnd.ms-excel;charset=${charset}`);
        window.navigator.msSaveBlob(blob, `${nmArquivo}.xls`);
    } else {
        link.href = uri + base64(format(template, ctx));
        link.download = `${nmArquivo}.xls`;
        link.click();
    }
};