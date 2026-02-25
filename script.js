// URL do Google Apps Script (Configurada)
const API_URL = "https://script.google.com/macros/s/AKfycby2lcmhw-TQOp4ehPTlA0MwgbGkusOZfPe-YkvF_3Qz3Ot8D8jldSKwESd0TnrG7zCRPQ/exec";

// --- ESTADO LOCAL (CACHE) ---
let clientes = [];
let transacoes = [];
let carregando = false;

// --- INICIALIZAÇÃO ---
document.addEventListener('DOMContentLoaded', () => {
    // verificarURL(); // Não precisa mais verificar, URL já está fixa
    carregarDados();
    carregarLogoSalva();

    // Configurar Upload de Logo
    document.getElementById('uploadLogo').addEventListener('change', function (e) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function (readerEvent) {
                const base64 = readerEvent.target.result;
                localStorage.setItem('logo_fiado_base64', base64);
                atualizarVisualLogo(base64);
            }
            reader.readAsDataURL(file);
        }
    });
});

function carregarLogoSalva() {
    const logoBase64 = localStorage.getItem('logo_fiado_base64');
    if (logoBase64) {
        atualizarVisualLogo(logoBase64);
    }
}

function atualizarVisualLogo(base64) {
    // Atualiza a imagem do topo
    document.getElementById('logoImg').src = base64;

    // Atualiza a Marca D'água (body::before) via CSS dinâmico
    let styleElem = document.getElementById('dynamic-watermark');
    if (!styleElem) {
        styleElem = document.createElement('style');
        styleElem.id = 'dynamic-watermark';
        document.head.appendChild(styleElem);
    }
    styleElem.innerHTML = `body::before { background-image: url('${base64}') !important; }`;
}

// --- FUNÇÕES DE API (FETCH) ---

async function carregarDados() {
    setCarregando(true);
    try {
        // Adiciona timestamp para evitar cache do navegador
        const url = `${API_URL}?op=ler&t=${Date.now()}`;
        const resp = await fetch(url);
        const json = await resp.json();

        if (json.clientes) {
            // Filtrar cabeçalhos (caso a planilha leia a linha 1 como dado)
            clientes = json.clientes.filter(c => c.id && c.id !== "ID" && c.nome !== "Nome");
        }
        if (json.transacoes) {
            transacoes = json.transacoes.filter(t => t.id !== "ID");
        }

        renderizarTudo();
    } catch (erro) {
        console.error(erro);
        alert("Erro ao carregar dados. Verifique sua conexão ou a URL do Script.");
    } finally {
        setCarregando(false);
    }
}

async function receberPagamento() {
    const clienteId = document.getElementById('detalheClienteId').value;
    const valorPagamento = parseFloat(document.getElementById('valorPagamento').value);

    if (!clienteId || isNaN(valorPagamento) || valorPagamento <= 0) {
        return alert("Selecione um cliente e digite um valor válido para pagamento.");
    }

    setCarregando(true);
    try {
        // Envia como transação normal, mas vamos tratar visualmente como pagamento
        // Sendo pagamento, entra como crédito (reduz a dívida).
        // A lógica de saldo é soma de (despesas) - (pagamentos).
        // Aqui vamos enviar valor NEGATIVO para abater.

        const resp = await fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify({
                op: 'lancar',
                clienteId: clienteId.toString(),
                descricao: "PAGAMENTO / ABATIMENTO",
                valor: -Math.abs(valorPagamento) // Garante negativo
            })
        });

        const json = await resp.json();

        if (json.sucesso) {
            transacoes.push({
                id: json.id,
                clienteId: clienteId.toString(),
                data: new Date().toLocaleDateString('pt-BR'),
                descricao: "PAGAMENTO / ABATIMENTO",
                valor: -Math.abs(valorPagamento)
            });

            document.getElementById('valorPagamento').value = '';
            alert("Pagamento registrado! Dívida abatida.");
            renderizarTudo();
            carregarDetalhesCliente(clienteId); // Atualiza a vista detalhada
        } else {
            alert("Erro ao registrar pagamento: " + json.erro);
        }
    } catch (e) {
        alert("Erro de conexão.");
    } finally {
        setCarregando(false);
    }
}

async function cadastrarCliente() {
    const nomeInput = document.getElementById('novoClienteNome');
    const nome = nomeInput.value.trim();

    if (!nome) return alert("Digite o nome do cliente!");

    setCarregando(true);
    try {
        const resp = await fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify({ op: 'criarCliente', nome: nome })
        });
        const json = await resp.json();

        if (json.sucesso) {
            // Atualização otimista LOCAL (para parecer rápido)
            clientes.push({ id: json.id, nome: nome });
            nomeInput.value = '';
            renderizarTudo();
            alert(`Cliente ${nome} cadastrado com sucesso!`);
        } else {
            alert("Erro ao salvar cliente: " + json.erro);
        }
    } catch (e) {
        alert("Erro de conexão ao salvar cliente.");
    } finally {
        setCarregando(false);
        // Recarregar para garantir sincronia (opcional)
        // carregarDados(); 
    }
}

async function lancarFiado() {
    const clienteId = document.getElementById('clienteSelect').value;
    const desc = document.getElementById('descItem').value.trim();
    const valor = parseFloat(document.getElementById('valorItem').value);

    if (!clienteId || !desc || isNaN(valor)) return alert("Preencha todos os campos!");

    setCarregando(true);
    try {
        const resp = await fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify({
                op: 'lancar',
                clienteId: clienteId.toString(), // Sheets IDs são strings geralmente
                descricao: desc,
                valor: valor
            })
        });
        const json = await resp.json();

        if (json.sucesso) {
            // Atualização otimista
            transacoes.push({
                id: json.id,
                clienteId: clienteId.toString(),
                data: new Date().toLocaleDateString('pt-BR'),
                descricao: desc,
                valor: valor
            });

            document.getElementById('descItem').value = '';
            document.getElementById('valorItem').value = '';
            renderizarTudo();
            alert("Fiado lançado na planilha!");
        } else {
            alert("Erro ao lançar: " + json.erro);
        }
    } catch (e) {
        alert("Erro de conexão ao lançar.");
    } finally {
        setCarregando(false);
    }
}

async function pagarItem(transacaoId, clienteId) {
    if (!confirm("Tem certeza que vai dar baixa neste item? Será removido da planilha.")) return;

    setCarregando(true);
    try {
        const resp = await fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify({ op: 'pagar', id: transacaoId })
        });
        const json = await resp.json();

        if (json.sucesso) {
            transacoes = transacoes.filter(t => t.id !== transacaoId);
            verDetalhes(clienteId); // Atualiza modal
            renderizarTudo(); // Atualiza fundo
        } else {
            alert("Erro ao pagar: " + json.erro);
        }
    } catch (e) {
        alert("Erro de conexão ao pagar.");
    } finally {
        setCarregando(false);
    }
}

// --- FUNÇÕES DE VISUALIZAÇÃO (UI) ---

function setCarregando(estado) {
    carregando = estado;
    const btns = document.querySelectorAll('button');
    btns.forEach(b => b.disabled = estado);
    document.body.style.cursor = estado ? 'wait' : 'default';

    // Opcional: Mostrar indicador visual
    const header = document.querySelector('header p');
    if (header) header.textContent = estado ? "Carregando dados da Planilha..." : "Sistema de Controle de Fiado";
}

// --- REGRAS DE NEGÓCIO ---

function renderizarTudo() {
    atualizarSelectClientes();
    // Renderiza a lista clean inicialmente para garantir que está pronta
    filtrarClientesClean();
}

function getDividaTotal(clienteId) {
    // Garantir comparação de strings para IDs
    return transacoes
        .filter(t => String(t.clienteId) === String(clienteId))
        .reduce((acc, curr) => acc + curr.valor, 0);
}

function atualizarSelectClientes() {
    const select = document.getElementById('clienteSelect');
    const valorAtual = select.value;

    select.innerHTML = '<option value="">Selecione o Cliente...</option>';

    clientes.forEach(c => {
        const option = document.createElement('option');
        option.value = c.id;
        option.textContent = c.nome;
        select.appendChild(option);
    });

    // Tenta manter seleção
    if (valorAtual) select.value = valorAtual;
}

// --- SISTEMA DE ABAS E BUSCA ---
let abaAtual = 'balcao';

function mostrarAba(aba) {
    abaAtual = aba;
    // Esconde tudo
    document.querySelectorAll('.tab-content').forEach(el => el.classList.remove('active'));
    document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

    // Mostra ativo
    document.getElementById(`tab-${aba}`).classList.add('active');
    document.getElementById(`btn-${aba}`).classList.add('active');

    if (aba === 'clientes') {
        filtrarClientesClean(); // Garante renderização ao entrar
        document.getElementById('buscaCliente').focus();
    }
}

function filtrarClientesClean() {
    const termo = document.getElementById('buscaCliente').value.toLowerCase();
    const tbody = document.getElementById('listaClientesCleanBody');
    tbody.innerHTML = '';

    // Ordena A-Z
    const clientesOrdenados = [...clientes].sort((a, b) => a.nome.localeCompare(b.nome));

    // Filtra
    const filtrados = clientesOrdenados.filter(c => c.nome.toLowerCase().includes(termo));

    if (filtrados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="2" style="text-align:center; color:#888;">Nenhum cliente encontrado.</td></tr>';
        return;
    }

    filtrados.forEach(c => {
        const tr = document.createElement('tr');
        // SEM COLUNA DE VALORES - Tabela Limpa
        tr.innerHTML = `
            <td style="font-size:1.1rem; padding: 15px;">${c.nome}</td>
            <td style="text-align:right;">
                <button class="btn-info" onclick="carregarDetalhesCliente('${c.id}')">Ver Ficha</button>
            </td>
        `;
        tbody.appendChild(tr);
    });
}

// Evento de Busca
document.getElementById('buscaCliente').addEventListener('keyup', filtrarClientesClean);

// Nova função de visualização embutida
function carregarDetalhesCliente(clienteId) {
    const cliente = clientes.find(c => String(c.id) === String(clienteId));
    if (!cliente) return;

    // Remove destaque de outros
    document.querySelectorAll('#listaClientesBody tr').forEach(tr => tr.style.backgroundColor = '');
    // Tenta destacar a linha (opcional, requer lógica extra pra achar o TR)

    // Configura painel
    document.getElementById('detalheClienteId').value = clienteId;
    document.getElementById('nomeDetalhe').textContent = cliente.nome;
    document.getElementById('painelDetalhes').style.display = 'block';

    // Rola até o painel
    document.getElementById('painelDetalhes').scrollIntoView({ behavior: 'smooth' });

    const lista = document.getElementById('listaTransacoesDetalhe');
    lista.innerHTML = '';

    const historico = transacoes.filter(t => String(t.clienteId) === String(clienteId));
    let total = 0;

    if (historico.length === 0) {
        lista.innerHTML = '<p style="color:#aaa; text-align:center;">Nenhum registro ainda.</p>';
    } else {
        const table = document.createElement('table');
        table.innerHTML = `
            <thead>
                <tr>
                    <th>Data</th>
                    <th>Descrição</th>
                    <th>Valor</th>
                    <th>Ação</th>
                </tr>
            </thead>
            <tbody></tbody>
        `;

        const tbody = table.querySelector('tbody');

        historico.forEach(t => {
            total += t.valor;
            const isPagamento = t.valor < 0;
            const style = isPagamento ? 'color: #27ae60; font-weight:bold;' : 'color: #c0392b;';
            const sinal = isPagamento ? '' : ''; // Negativo já tem sinal

            const tr = document.createElement('tr');
            tr.innerHTML = `
                <td>${t.data}</td>
                <td>${t.descricao}</td>
                <td style="${style}">R$ ${t.valor.toFixed(2)}</td>
                <td><button onclick="removerItem('${t.id}', '${clienteId}')" class="btn-danger" style="padding:2px 8px; font-size:0.7rem">Excluir</button></td>
            `;
            tbody.appendChild(tr);
        });

        lista.appendChild(table);
    }

    // Atualiza saldo visual
    const saldoElem = document.getElementById('saldoTotalDetalhe');
    saldoElem.textContent = total.toFixed(2);
    saldoElem.style.color = total > 0 ? '#e74c3c' : '#27ae60'; // Vermelho se deve, Verde se tem crédito
}

async function removerItem(transacaoId, clienteId) {
    if (!confirm("Apagar este registro? Se for erro de lançamento ok, mas para pagamento use o campo de 'Receber Pagamento'.")) return;

    setCarregando(true);
    try {
        const resp = await fetch(API_URL, {
            method: 'POST',
            body: JSON.stringify({ op: 'pagar', id: transacaoId }) // Mantivemos o nome 'pagar' no backend para apagar
        });
        const json = await resp.json();

        if (json.sucesso) {
            transacoes = transacoes.filter(t => t.id !== transacaoId);
            renderizarTudo();
            carregarDetalhesCliente(clienteId);
        } else {
            alert("Erro ao excluir: " + json.erro);
        }
    } catch (e) {
        alert("Erro de conexão.");
    } finally {
        setCarregando(false);
    }
}
// Remover funções antigas de modal
function verDetalhes() { }
function fecharModal() { }
function pagarItem() { }
