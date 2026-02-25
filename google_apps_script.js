// --- CONFIGURAÇÃO ---
const ID_PLANILHA = SpreadsheetApp.getActiveSpreadsheet().getId();
const CACHE_MINUTOS = 0; // Cache desativado para garantir tempo real na troca de dados

function doGet(e) {
    const op = e.parameter.op;
    const ultimoTimestampCliente = e.parameter.ts || 0;

    if (op === "ler") return lerDados(ultimoTimestampCliente);

    return ContentService.createTextOutput("Operação inválida.");
}

function doPost(e) {
    // Bloqueio para evitar conflitos de escrita simultânea
    const lock = LockService.getScriptLock();
    // Tenta esperar até 10s. Se não conseguir, erro.
    if (!lock.tryLock(10000)) {
        return respostaJSON({ erro: "Servidor ocupado. Tente novamente." });
    }

    try {
        const dados = JSON.parse(e.postData.contents);
        const op = dados.op;

        let resultado;

        // Roteamento de Operações
        if (op === "criarCliente") resultado = criarCliente(dados.nome, dados.telefone || "");
        else if (op === "lancar") resultado = lancarTransacao(dados.clienteId, dados.descricao, dados.valor);
        else if (op === "abater") resultado = abaterDivida(dados.clienteId, dados.valor);
        else if (op === "pagar") resultado = pagarTransacao(dados.id);
        else if (op === "excluirCliente") resultado = excluirCliente(dados.clienteId);
        else if (op === "editarCliente") resultado = editarCliente(dados.clienteId, dados.novoNome, dados.novoTelefone);
        else if (op === "pagarLote") resultado = pagarLote(dados.ids);
        else if (op === "registrarFluxo") resultado = registrarFluxo(dados.tipo, dados.descricao, dados.valor, dados.metodo);
        else if (op === "excluirEntrada") resultado = excluirMovimentacao(dados.id, "Entradas");
        else if (op === "excluirSaida") resultado = excluirMovimentacao(dados.id, "Saidas");
        else resultado = { erro: "Operação desconhecida" };

        // Se a operação deu certo (e não foi só leitura), atualizamos o Timestamp Global
        if (resultado && resultado.sucesso) {
            marcarAlteracao();
        }

        return respostaJSON(resultado);

    } catch (err) {
        return respostaJSON({ erro: err.toString() });
    } finally {
        lock.releaseLock();
    }
}

// --- CONTROLE DE VERSÃO (OTIMIZAÇÃO) ---

function marcarAlteracao() {
    // Salva o momento atual (milissegundos) como a última vez que algo mudou
    const agora = Date.now().toString();
    PropertiesService.getScriptProperties().setProperty('ULTIMA_MODIFICACAO', agora);
}

function lerDados(ultimoTimestampCliente) {
    const props = PropertiesService.getScriptProperties();
    // Pega a última vez que o servidor mudou. Se não tiver, define como atual.
    let ultimaModifServidor = props.getProperty('ULTIMA_MODIFICACAO');

    if (!ultimaModifServidor) {
        ultimaModifServidor = Date.now().toString();
        props.setProperty('ULTIMA_MODIFICACAO', ultimaModifServidor);
    }

    // COMPARAÇÃO MAGIC:
    // Se o cliente diz que tem a versão X, e o servidor ainda está na versão X,
    // não precisamos enviar dados nenhuns! Economiza muito tempo.
    if (ultimoTimestampCliente && String(ultimoTimestampCliente) === String(ultimaModifServidor)) {
        return respostaJSON({
            sucesso: true,
            modificado: false,
            timestamp: ultimaModifServidor
        });
    }

    // Se chegou aqui, é porque TEM dados novos ou o cliente está zerado.
    // LÊ A PLANILHA (Parte Lenta)
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheetClientes = verificarEstrutura(ss, "Clientes", ["ID", "Nome", "Telefone"]);
    const sheetTransacoes = verificarEstrutura(ss, "Transacoes", ["ID", "ClienteID", "Data", "Descricao", "Valor"]);
    const sheetEntradas = verificarEstrutura(ss, "Entradas", ["ID", "Data", "Descricao", "Valor", "Metodo"]);
    const sheetSaidas = verificarEstrutura(ss, "Saidas", ["ID", "Data", "Descricao", "Valor"]);

    const dadosClientes = sheetClientes.getDataRange().getValues().slice(1);
    const listaClientes = dadosClientes.filter(r => r[0] !== "").map(row => ({
        id: row[0],
        nome: row[1],
        telefone: row[2]
    }));

    const dadosTransacoes = sheetTransacoes.getDataRange().getValues().slice(1);
    const listaTransacoes = dadosTransacoes.filter(r => r[0] !== "").map(row => ({
        id: row[0],
        clienteId: row[1],
        data: formatarData(row[2]),
        descricao: row[3],
        valor: Number(row[4])
    }));

    const dadosEntradas = sheetEntradas.getDataRange().getValues().slice(1);
    const listaEntradas = dadosEntradas.filter(r => r[0] !== "").map(row => ({
        id: row[0],
        data: formatarData(row[1]),
        descricao: row[2],
        valor: Number(row[3]),
        metodo: row[4]
    }));

    const dadosSaidas = sheetSaidas.getDataRange().getValues().slice(1);
    const listaSaidas = dadosSaidas.filter(r => r[0] !== "").map(row => ({
        id: row[0],
        data: formatarData(row[1]),
        descricao: row[2],
        valor: Number(row[3])
    }));

    return respostaJSON({
        sucesso: true,
        modificado: true,
        timestamp: ultimaModifServidor,
        clientes: listaClientes,
        transacoes: listaTransacoes,
        entradas: listaEntradas,
        saidas: listaSaidas
    });
}

// --- FUNÇÕES DE ESCRITA (Mantidas, mas limpas para caber) ---

function criarCliente(nome, telefone) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = verificarEstrutura(ss, "Clientes", ["ID", "Nome", "Telefone"]);
    let novoId = 1;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
        const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues().flat();
        const idsNum = ids.map(x => parseInt(x)).filter(x => !isNaN(x));
        if (idsNum.length > 0) novoId = Math.max(...idsNum) + 1;
    }
    sheet.appendRow([novoId, nome, telefone]);
    return { sucesso: true, id: novoId };
}

function lancarTransacao(clienteId, descricao, valor) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = verificarEstrutura(ss, "Transacoes", ["ID", "ClienteID", "Data", "Descricao", "Valor"]);
    const novoId = Date.now().toString();
    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    sheet.appendRow([novoId, clienteId, dataHoje, descricao, valor]);
    return { sucesso: true, id: novoId };
}

function abaterDivida(clienteId, valorPagamento) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("Transacoes");
    if (!sheet) return { erro: "Aba Transacoes não encontrada" };

    let valorRestante = Number(valorPagamento);
    if (isNaN(valorRestante) || valorRestante <= 0) return { erro: "Valor inválido" };

    // Otimização: Ler tudo uma vez
    const range = sheet.getDataRange();
    const values = range.getValues();

    let dividas = [];
    // Index 4 = Valor, Index 1 = ClienteID
    for (let i = 1; i < values.length; i++) {
        if (String(values[i][1]) === String(clienteId) && Number(values[i][4]) > 0) {
            dividas.push({ rowIndex: i + 1, valor: Number(values[i][4]) });
        }
    }

    let deletar = [];
    let atualizar = [];

    for (let divida of dividas) {
        if (valorRestante <= 0) break;
        let p = Math.min(divida.valor, valorRestante);
        let novoSaldo = divida.valor - p;
        if (novoSaldo <= 0.01) deletar.push(divida.rowIndex);
        else atualizar.push({ rowIndex: divida.rowIndex, val: novoSaldo });
        valorRestante -= p;
    }

    atualizar.forEach(o => sheet.getRange(o.rowIndex, 5).setValue(o.val));
    deletar.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));

    if (valorRestante > 0.01) {
        sheet.appendRow([Date.now(), clienteId, Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy"), "CRÉDITO/TROCO", -valorRestante]);
    }

    return { sucesso: true, abatido: valorPagamento - valorRestante, credito: valorRestante };
}

function pagarTransacao(id) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("Transacoes");
    const dados = sheet.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(id)) {
            sheet.deleteRow(i + 1);
            return { sucesso: true };
        }
    }
    return { erro: "Não encontrado" };
}

function excluirCliente(id) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sTr = ss.getSheetByName("Transacoes");
    if (sTr) {
        let d = sTr.getDataRange().getValues();
        let dels = [];
        for (let i = 1; i < d.length; i++) if (String(d[i][1]) === String(id)) dels.push(i + 1);
        dels.sort((a, b) => b - a).forEach(r => sTr.deleteRow(r));
    }
    const sCli = ss.getSheetByName("Clientes");
    let dC = sCli.getDataRange().getValues();
    for (let i = 1; i < dC.length; i++) {
        if (String(dC[i][0]) === String(id)) {
            sCli.deleteRow(i + 1);
            return { sucesso: true };
        }
    }
    return { erro: "Cliente não achado" };
}

function editarCliente(id, nome, tel) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("Clientes");
    const d = sheet.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
        if (String(d[i][0]) === String(id)) {
            if (nome) sheet.getRange(i + 1, 2).setValue(nome);
            if (tel !== undefined) sheet.getRange(i + 1, 3).setValue(tel);
            return { sucesso: true };
        }
    }
    return { erro: "Cliente não achado" };
}

function pagarLote(ids) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName("Transacoes");
    const d = sheet.getDataRange().getValues();
    const header = d[0];
    const rows = d.slice(1);

    // Filtro Inverso: Manter o que NÃO está na lista de IDs
    const novos = rows.filter(r => !ids.includes(String(r[0])));

    if (novos.length === rows.length) return { sucesso: true, msg: "Nada mudou" };

    sheet.clearContents();
    const final = [header, ...novos];
    if (final.length > 0) sheet.getRange(1, 1, final.length, final[0].length).setValues(final);

    return { sucesso: true };
}

function registrarFluxo(tipo, descricao, valor, metodo) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const dataHoje = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy");
    const novoId = Date.now().toString();

    if (tipo === "entrada") {
        const sheet = verificarEstrutura(ss, "Entradas", ["ID", "Data", "Descricao", "Valor", "Metodo"]);
        sheet.appendRow([novoId, dataHoje, descricao, valor, metodo]);
    } else {
        const sheet = verificarEstrutura(ss, "Saidas", ["ID", "Data", "Descricao", "Valor"]);
        sheet.appendRow([novoId, dataHoje, descricao, valor]);
    }

    return { sucesso: true };
}

function excluirMovimentacao(id, aba) {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const sheet = ss.getSheetByName(aba);
    if (!sheet) return { erro: "Aba " + aba + " não encontrada" };
    const dados = sheet.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][0]) === String(id)) {
            sheet.deleteRow(i + 1);
            return { sucesso: true };
        }
    }
    return { erro: "Lançamento não encontrado" };
}

// --- UTILITÁRIOS ---

function verificarEstrutura(ss, nome, headers) {
    let sheet = ss.getSheetByName(nome);
    if (!sheet) {
        sheet = ss.insertSheet(nome);
        sheet.appendRow(headers);
    }
    return sheet;
}

function formatarColunas(sheet, tipo) { /* Mantém formatação manual ou simples */ }

function formatarData(d) {
    if (d instanceof Date) return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    return d;
}

function respostaJSON(obj) {
    return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
