
function verifyJQuery(){
    if (typeof jQuery == 'undefined') {
        var script = document.createElement('script');
        script.src = 'http://code.jquery.com/jquery-1.11.0.min.js';
        script.type = 'text/javascript';
        document.getElementsByTagName('head')[0].appendChild(script);
    }
}

function findDataByContent(elements, content){
    return elements
        .find("p:contains(\'" + content + "\')")[0]
        .innerText.replace(content, '')
        .trim().toUpperCase();
}

function downloadFile(fileName, urlData) {
    var aLink = document.createElement('a');
    aLink.download = fileName;
    aLink.href = urlData;
    var event = new MouseEvent('click');
    aLink.dispatchEvent(event);
}

function createDescription(user, resp, coresp){
    var desc;
    if (resp == '-') {
        desc = 'Nenhuma informação de patrimônio.';
    } else if (resp == user) {
        desc = 'O responsável patrimonial é o usuário da máquina.';
    } else if (coresp == user) {
        desc = 'O co-responsável patrimonial é o usuário da máquina.';
    } else {
        desc = 'O usuário não está vinculado ao patrimônio.';
    }
    return desc;
}

function loadDetailPage(ips, index){
    var url = window.location.href;
    var block_id = url.substring(url.indexOf('=') + 1);
    $('#crawler-iframe').attr('src', 'https://www1.ufrgs.br/RegistroEstacoes/Operacoes/ipdetails.php?IP=' + ips[index] + '&blocoConsulta=' + block_id);
}

function executeCrawling(){
    verifyJQuery();
    var index = 0;
    var ipv4List = []
    var data = [];

    $('#quadrobranco-dir').append('<iframe id="crawler-iframe"> </iframe>');
    $('#crawler-iframe').load(function(){
        if ($(this).attr('src').match(/https:\/\/www1.ufrgs.br\/RegistroEstacoes\/Operacoes\/ipdetails.php/g)){
            
            console.log('Verificando ' + ipv4List[index] + ':');
            var elements = $(this).contents().find('.fieldset-1');
            var user = findDataByContent(elements, 'Nome do Usuário:');
            if (elements.find("p:contains('Responsável:')").length > 0) {
                var resp = findDataByContent(elements, 'Responsável:');
                var coresp = findDataByContent(elements, 'Co-Responsável:');
                var machineId = findDataByContent(elements, 'Patrimônio:');
            } else {
                var resp = '-';
                var coresp = '-';
                var machineId = '-';
            }
            var desc = createDescription(user, resp, coresp)

            console.log('- ' + desc);
            data.push(ipv4List[index] + ';' + machineId + ';' + user + ';' + resp + ';' + coresp + ';' + desc)
            index++;

            if (index < ipv4List.length) {
                loadDetailPage(ipv4List, index);
            } else {
                var csv = 'IPV4;PATRIMÔNIO;USUÁRIO DO NAC;RESPONSÁVEL PATRIMONIAL;CO-RESPONSÁVEL PATRIMONIAL;ESTADO\n';
                for(var i in data) {
                    csv += data[i] + '\n';
                }
                var file_name = prompt("Como deseja nomear o arquivo?", "relatorio-nac-patrimonio");
                if (file_name == null || file_name == "") {
                    file_name = "relatorio-nac-patrimonio";
                }
                downloadFile(file_name + '.csv', 'data:attachment/csv;charset=utf-8,%EF%BB%BF' + encodeURIComponent(csv));
            }
        }
    });

    $('.usado').each(function(){
        var text = $(this).find('a').attr('href');
        var firstIndex = text.indexOf('\'') + 1;
        var lastIndex = text.substr(firstIndex).indexOf('\'');
        ipv4List.push(text.substr(firstIndex, lastIndex));
    });
    
    loadDetailPage(ipv4List, index);
}

executeCrawling();