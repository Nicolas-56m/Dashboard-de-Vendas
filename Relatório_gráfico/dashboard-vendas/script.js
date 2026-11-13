let dadosOriginais = [];
let charts = {};

const btnImportar = document.getElementById('btn-importar');
const inputExcel = document.getElementById('input-excel');
const filtroRegiao = document.getElementById('filtro-regiao');
const filtroProduto = document.getElementById('filtro-produto');
const btnLimpar = document.getElementById('btn-limpar');

btnImportar.addEventListener('click', () => inputExcel.click());

inputExcel.addEventListener('change', (evento) => {
    const arquivo = evento.target.files[0];
    if (!arquivo) return;

    const reader = new FileReader();
    reader.onload = (e) => {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            const primeiraAba = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[primeiraAba];

            // Pega os dados como objetos normais usando a primeira linha como cabeçalho
            const linhasBrutas = XLSX.utils.sheet_to_json(worksheet, { defval: "" });

            dadosOriginais = [];

            linhasBrutas.forEach(linha => {
                // Converte as chaves do objeto para minúsculo para localizar sem erro de acento ou maiúscula
                const chaves = Object.keys(linha);
                
                const acharValor = (termos) => {
                    for (let termo of termos) {
                        const chaveReal = chaves.find(k => k.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").includes(termo));
                        if (chaveReal && linha[chaveReal] !== undefined) {
                            return linha[chaveReal];
                        }
                    }
                    return '';
                };

                const regiao = String(acharValor(['regiao'])).trim();
                const produto = String(acharValor(['produto', 'produtos'])).trim();
                const total = limparMoeda(acharValor(['total']));

                // Só adiciona se tiver produto ou região válida e não for a linha de rodapé "Total"
                if ((regiao !== '' || produto !== '') && !regiao.toLowerCase().includes('total')) {
                    dadosOriginais.push({
                        regiao: regiao || 'Outros',
                        estado: String(acharValor(['estado'])).trim(),
                        produto: produto || 'Outros',
                        data: formatarDataExcel(acharValor(['data'])),
                        quantidade: Number(acharValor(['quantidade', 'quant'])) || 0,
                        valorUnitario: limparMoeda(acharValor(['unitario', 'valor unitario'])),
                        frete: limparMoeda(acharValor(['frete'])),
                        total: total
                    });
                }
            });

            if (dadosOriginais.length === 0) {
                alert("Nenhum dado válido foi encontrado na planilha. Verifique se o arquivo está correto.");
                return;
            }

            preencherFiltros();
            atualizarDashboard();
        } catch (erro) {
            alert("Erro ao processar o arquivo Excel: " + erro.message);
        }
    };
    reader.readAsArrayBuffer(arquivo);
});

function limparMoeda(valor) {
    if (typeof valor === 'number') return valor;
    if (!valor) return 0;
    const limpo = String(valor)
        .replace('R$', '')
        .replace(/\./g, '')
        .replace(',', '.')
        .trim();
    return parseFloat(limpo) || 0;
}

function formatarDataExcel(valor) {
    if (typeof valor === 'number') {
        const data = new Date(Math.round((valor - 25569) * 86400 * 1000));
        return data.toLocaleDateString('pt-BR', { timeZone: 'UTC' });
    }
    return String(valor).trim();
}

function formatarBRL(valor) {
    return valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
}

function preencherFiltros() {
    const regioes = [...new Set(dadosOriginais.map(d => d.regiao))].sort();
    const produtos = [...new Set(dadosOriginais.map(d => d.produto))].sort();

    filtroRegiao.innerHTML = '<option value="">Todas as regiões</option>';
    regioes.forEach(r => {
        filtroRegiao.innerHTML += `<option value="${r}">${r}</option>`;
    });

    filtroProduto.innerHTML = '<option value="">Todos os produtos</option>';
    produtos.forEach(p => {
        filtroProduto.innerHTML += `<option value="${p}">${p}</option>`;
    });
}

filtroRegiao.addEventListener('change', atualizarDashboard);
filtroProduto.addEventListener('change', atualizarDashboard);
btnLimpar.addEventListener('click', () => {
    filtroRegiao.value = '';
    filtroProduto.value = '';
    atualizarDashboard();
});

function atualizarDashboard() {
    const regiaoSel = filtroRegiao.value;
    const produtoSel = filtroProduto.value;

    const filtrados = dadosOriginais.filter(item => {
        const bateRegiao = !regiaoSel || item.regiao === regiaoSel;
        const bateProduto = !produtoSel || item.produto === produtoSel;
        return bateRegiao && bateProduto;
    });

    const faturamentoTotal = filtrados.reduce((acc, item) => acc + item.total, 0);
    const unidadesTotal = filtrados.reduce((acc, item) => acc + item.quantidade, 0);
    const freteTotal = filtrados.reduce((acc, item) => acc + item.frete, 0);

    document.getElementById('kpi-faturamento').textContent = formatarBRL(faturamentoTotal);
    document.getElementById('kpi-unidades').textContent = unidadesTotal;
    document.getElementById('kpi-pedidos').textContent = filtrados.length;
    document.getElementById('kpi-frete').textContent = formatarBRL(freteTotal);

    atualizarTabela(filtrados);
    atualizarGraficos(filtrados);
}

function atualizarTabela(dados) {
    const tbody = document.getElementById('tabela-corpo');
    if (dados.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8">Nenhum registro encontrado.</td></tr>';
        return;
    }

    tbody.innerHTML = dados.map(d => `
        <tr>
            <td>${d.regiao}</td>
            <td>${d.estado}</td>
            <td>${d.produto}</td>
            <td>${d.data}</td>
            <td>${d.quantidade}</td>
            <td>${formatarBRL(d.valorUnitario)}</td>
            <td>${formatarBRL(d.frete)}</td>
            <td><strong>${formatarBRL(d.total)}</strong></td>
        </tr>
    `).join('');
}

function atualizarGraficos(dados) {
    const porRegiao = agruparSoma(dados, 'regiao', 'total');
    const porProduto = agruparSoma(dados, 'produto', 'quantidade');
    const porData = agruparSoma(dados, 'data', 'total');
    const freteRegiao = agruparSoma(dados, 'regiao', 'frete');

    renderizarGrafico('grafico-regiao', 'bar', Object.keys(porRegiao), Object.values(porRegiao), 'Faturamento (R$)', '#2563eb');
    renderizarGrafico('grafico-produto', 'doughnut', Object.keys(porProduto), Object.values(porProduto), 'Unidades Vendidas', ['#2563eb', '#0ea5e9', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444']);
    renderizarGrafico('grafico-evolucao', 'line', Object.keys(porData), Object.values(porData), 'Vendas por Data (R$)', '#10b981');
    renderizarGrafico('grafico-frete', 'bar', Object.keys(freteRegiao), Object.values(freteRegiao), 'Frete (R$)', '#f59e0b');
}

function agruparSoma(lista, chaveGrupo, chaveValor) {
    return lista.reduce((acc, item) => {
        acc[item[chaveGrupo]] = (acc[item[chaveGrupo]] || 0) + item[chaveValor];
        return acc;
    }, {});
}

function renderizarGrafico(idCanvas, tipo, labels, valores, labelDataset, cores) {
    if (charts[idCanvas]) {
        charts[idCanvas].destroy();
    }

    const ctx = document.getElementById(idCanvas).getContext('2d');
    charts[idCanvas] = new Chart(ctx, {
        type: tipo,
        data: {
            labels: labels,
            datasets: [{
                label: labelDataset,
                data: valores,
                backgroundColor: cores,
                borderColor: tipo === 'line' ? cores : '#ffffff',
                borderWidth: 2,
                tension: 0.3
            }]
        },
        options: {
            responsive: true,
            plugins: {
                legend: { display: tipo === 'doughnut' }
            }
        }
    });
}