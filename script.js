var dados = [];

// Função para limpar os campos do formulário
function limparCampos() {
    document.getElementById('nomeAnalista').value = '';
    document.getElementById('loginCliente').value = '';
    document.getElementById('olt').value = '';
    document.getElementById('vlan').value = '';
    document.getElementById('serialEquipamento').value = '';
    document.getElementById('macAddress').value = '';
    document.getElementById('potencia').value = '';
    document.getElementById('acessoRemoto').value = '';
}

document.getElementById('formulario').addEventListener('submit', function(event) {
    event.preventDefault(); // Evita o envio padrão do formulário
    
    // Captura dos valores do formulário
    var nomeAnalista = document.getElementById('nomeAnalista').value;
    var loginCliente = document.getElementById('loginCliente').value;
    var olt = document.getElementById('olt').value;
    var vlan = document.getElementById('vlan').value;
    var serialEquipamento = document.getElementById('serialEquipamento').value;
    var macAddress = document.getElementById('macAddress').value;
    var potencia = document.getElementById('potencia').value;
    var acessoRemoto = document.getElementById('acessoRemoto').value;
    
    // Adiciona os dados ao array
    dados.push([nomeAnalista, loginCliente, olt, vlan, serialEquipamento, macAddress, potencia, acessoRemoto]);
});

document.getElementById('salvarPlanilha').addEventListener('click', function() {
    // Criação de uma nova planilha Excel
    var wb = XLSX.utils.book_new();
    var ws = XLSX.utils.aoa_to_sheet([['Nome do Analista', 'Login do Cliente', 'OLT', 'VLAN', 'Serial do Equipamento', 'Mac-Address', 'Potencia', 'Acesso Remoto']].concat(dados));

    // Conversão para o formato de arquivo Excel
    var wbout = XLSX.write(wb, { bookType: 'xlsx', type: 'binary' });
    function s2ab(s) {
        var buf = new ArrayBuffer(s.length);
        var view = new Uint8Array(buf);
        for (var i = 0; i < s.length; i++) view[i] = s.charCodeAt(i) & 0xFF;
        return buf;
    }
    
    // Download do arquivo Excel
    var blob = new Blob([s2ab(wbout)], { type: 'application/octet-stream' });
    var link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = 'planilha.xlsx';
    link.click();

    // Limpar os campos após salvar a planilha
    limparCampos();
});
