import { useState, useMemo, useRef, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { Upload, TrendingUp, Calculator, FileSpreadsheet, LineChart, Package } from 'lucide-react';
import {
  calcDasHoje,
  calcDasPorDentroRate,
  calcDasPorForaRate,
  REFORM_SCHEDULE,
  CATEGORIAS_CREDITO,
  fmtPct,
  fmtR,
} from './simplesEngine';
import produtosBaseUrl from './assets/produtos-base.xlsx?url';
import {
  gerarId, custoTotal, salvarSnapshotLocal, carregarOuSemear,
  criarProduto, excluirProduto, atualizarCustoBase, criarCustoExtra, excluirCustoExtra, substituirTodosProdutos,
} from './produtosStore';

const ANOS_REFORMA = Object.keys(REFORM_SCHEDULE);

// ─── Inputs numéricos em BRL ────────────────────────────────────────────────
// Aceita tanto "8,8" (vírgula BR) quanto "8.8" (ponto, comum ao digitar
// alíquotas) e "1.234,56" (milhar BR) — vírgula, quando presente, é sempre o
// separador decimal; sem vírgula, o ponto é tratado como decimal.
const parseBR = (v) => {
  const cleaned = String(v ?? '').replace(/[^\d,.-]/g, '');
  if (!cleaned) return 0;
  const normalized = cleaned.includes(',') ? cleaned.replace(/\./g, '').replace(',', '.') : cleaned;
  const n = parseFloat(normalized);
  return Number.isFinite(n) ? n : 0;
};

function CurrencyInput({ label, value, onChange, placeholder }) {
  return (
    <label className="block">
      <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">{label}</span>
      <input
        type="text"
        inputMode="decimal"
        value={value}
        onChange={(e) => onChange(e.target.value)}
        placeholder={placeholder}
        className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none"
      />
    </label>
  );
}

// ─── Parser da planilha de produtos ────────────────────────────────────────
function parseProdutos(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: 'array' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: null });
  if (!rows.length) return [];
  const header = rows[0].map((h) => String(h || '').toLowerCase());
  const findCol = (needle, fallbackIdx) => {
    const idx = header.findIndex((h) => h.includes(needle));
    return idx >= 0 ? idx : fallbackIdx;
  };
  const skuCol = findCol(header.some((h) => h.includes('produto')) ? 'produto' : 'sku', 0);
  const vendaCol = findCol('venda', 1);
  const custoCol = findCol('custo', 7);

  const out = [];
  for (let i = 1; i < rows.length; i++) {
    const r = rows[i];
    if (!r || !r[skuCol]) continue;
    const venda = parseFloat(r[vendaCol]);
    const custo = parseFloat(r[custoCol]);
    if (!Number.isFinite(venda) || !Number.isFinite(custo) || venda <= 0) continue;
    out.push({ id: gerarId(), sku: String(r[skuCol]), venda, custo, custosExtras: [] });
  }
  return out;
}

// ─── Aba: Repasse (visual) ──────────────────────────────────────────────────
// A alíquota REAL (pelo RBT12 informado) costuma vir maior que a alíquota que
// os preços atuais já embutem (o negócio cresceu de faixa e o preço ainda não
// acompanhou). Em vez de repassar tudo de uma vez, distribui o aumento em
// etapas mensais, preservando a margem que os preços de hoje já consideram —
// não a margem real de hoje, que já está mais apertada do que parece.
function RepasseTab({ rbt12Raw, rbt12RealRaw, produtos, aliquotaUsadaRaw, setAliquotaUsadaRaw, mesesRaw, setMesesRaw }) {
  const rbt12Declarado = parseBR(rbt12Raw);
  const rbt12Real = parseBR(rbt12RealRaw);
  const rbt12 = rbt12Real > 0 ? rbt12Real : rbt12Declarado;
  const aliquotaReal = useMemo(() => (rbt12 > 0 ? calcDasHoje(rbt12).aliqEf : 0), [rbt12]);
  const aliquotaUsada = parseBR(aliquotaUsadaRaw) / 100;
  const meses = Math.max(1, parseInt(mesesRaw, 10) || 0);
  const gap = aliquotaReal - aliquotaUsada;

  const produtosValidos = useMemo(() => produtos.filter((p) => p.venda > 0), [produtos]);
  const margemAlvoMedia = useMemo(() => {
    if (!produtosValidos.length) return 0;
    const soma = produtosValidos.reduce((acc, p) => acc + (1 - custoTotal(p) / p.venda - aliquotaUsada), 0);
    return soma / produtosValidos.length;
  }, [produtosValidos, aliquotaUsada]);

  const pontos = useMemo(() => Array.from({ length: meses + 1 }, (_, mes) => {
    const aliquotaMes = aliquotaUsada + gap * (mes / meses);
    return { mes, aliquotaMes, margemComRepasse: margemAlvoMedia, margemSemRepasse: margemAlvoMedia - gap * (mes / meses) };
  }), [aliquotaUsada, gap, meses, margemAlvoMedia]);

  const [mesSelecionado, setMesSelecionado] = useState(meses);
  useEffect(() => { setMesSelecionado(meses); }, [meses]);
  const aliquotaSelecionada = aliquotaUsada + gap * (mesSelecionado / meses);

  const linhasMes = useMemo(() => produtosValidos.map((p) => {
    const custo = custoTotal(p);
    const margemAlvo = 1 - custo / p.venda - aliquotaUsada;
    const denom = 1 - margemAlvo - aliquotaSelecionada;
    const precoSugerido = denom > 0 ? custo / denom : null;
    const margemSemRepasse = 1 - custo / p.venda - aliquotaSelecionada;
    return { ...p, custo, margemAlvo, precoSugerido, margemSemRepasse };
  }), [produtosValidos, aliquotaUsada, aliquotaSelecionada]);

  const temDados = !!aliquotaUsadaRaw && rbt12 > 0;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-slate-700 mb-1">Repasse suave até a alíquota real</h3>
        <p className="text-[11px] text-slate-400 mb-4">
          Se os preços de hoje ainda consideram uma alíquota menor do que a real (a empresa cresceu de faixa),
          em vez de repassar tudo de uma vez, distribui o aumento em etapas mensais até a alíquota real.
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <CurrencyInput label="Alíquota já embutida nos preços hoje (%)" value={aliquotaUsadaRaw} onChange={setAliquotaUsadaRaw} placeholder="Ex: 6,00" />
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Meses até o repasse completo</span>
            <input
              type="number" min="1" value={mesesRaw} onChange={(e) => setMesesRaw(e.target.value)}
              className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none"
            />
          </label>
        </div>
        {rbt12 <= 0 && <p className="text-sm text-slate-400 italic mt-3">Informe o RBT12 na aba Simples Nacional primeiro.</p>}
        {rbt12 > 0 && (
          <p className="text-[11px] text-slate-400 mt-3">
            Alíquota real calculada com base no RBT12 {rbt12Real > 0 ? 'real' : 'declarado'} ({fmtR(rbt12)}) informado na aba Simples Nacional.
          </p>
        )}
      </div>

      {temDados && (
        <>
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <div className="flex flex-wrap items-baseline gap-x-6 gap-y-1 mb-4">
              <p className="text-sm text-slate-600">Hoje nos preços: <b>{fmtPct(aliquotaUsada)}</b></p>
              <p className="text-sm text-slate-600">Real (RBT12 informado): <b>{fmtPct(aliquotaReal, 4)}</b></p>
              <p className="text-sm text-slate-600">Gap: <b className={gap > 0 ? 'text-amber-600' : 'text-slate-500'}>{fmtPct(gap, 4)}</b></p>
              <p className="text-sm text-slate-600">Margem-alvo média: <b>{fmtPct(margemAlvoMedia)}</b></p>
            </div>
            {gap <= 0 ? (
              <p className="text-sm text-slate-400 italic">Sem gap a repassar — a alíquota real já é igual ou menor que a usada nos preços.</p>
            ) : produtosValidos.length === 0 ? (
              <p className="text-sm text-slate-400 italic">Carregue produtos na aba Produtos para ver a margem projetada.</p>
            ) : null}
          </div>

          {gap > 0 && produtosValidos.length > 0 && (
            <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
              <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">Escolha o mês</h4>
              <div className="flex flex-wrap gap-2 mb-5">
                {pontos.map((p) => (
                  <button
                    key={p.mes}
                    onClick={() => setMesSelecionado(p.mes)}
                    className={`px-3 py-1.5 rounded-lg text-sm font-semibold border transition ${
                      mesSelecionado === p.mes
                        ? 'bg-[#2F6F4E] text-white border-[#2F6F4E]'
                        : 'bg-white text-slate-600 border-slate-200 hover:border-[#2F6F4E]'
                    }`}
                  >
                    {p.mes === 0 ? 'Hoje' : `Mês ${p.mes}`}
                  </button>
                ))}
              </div>

              <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 mb-5">
                <div className="bg-slate-50 rounded-xl p-4">
                  <p className="text-xs font-semibold text-slate-500 uppercase">Alíquota do mês</p>
                  <p className="text-2xl font-black text-slate-800 mt-1">{fmtPct(aliquotaSelecionada, 4)}</p>
                </div>
                <div className="bg-emerald-50 rounded-xl p-4">
                  <p className="text-xs font-semibold text-emerald-700 uppercase">Margem atual — com repasse</p>
                  <p className="text-2xl font-black text-emerald-700 mt-1">{fmtPct(margemAlvoMedia)}</p>
                  <p className="text-[11px] text-emerald-600 mt-1">preservada, seguindo o cronograma</p>
                </div>
                <div className="bg-amber-50 rounded-xl p-4">
                  <p className="text-xs font-semibold text-amber-700 uppercase">Margem atual — sem repasse</p>
                  <p className="text-2xl font-black text-amber-700 mt-1">{fmtPct(pontos[mesSelecionado].margemSemRepasse)}</p>
                  <p className="text-[11px] text-amber-600 mt-1">se o preço não mudar até esse mês</p>
                </div>
              </div>

              <h4 className="text-xs font-bold text-slate-500 uppercase mb-2">Detalhe por produto — {mesSelecionado === 0 ? 'hoje' : `mês ${mesSelecionado}`}</h4>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="bg-slate-50 text-left text-[11px] uppercase text-slate-500">
                      <th className="px-4 py-2">SKU</th>
                      <th className="px-4 py-2 text-right">Preço atual</th>
                      <th className="px-4 py-2 text-right">Preço sugerido no mês</th>
                      <th className="px-4 py-2 text-right">Margem com repasse</th>
                      <th className="px-4 py-2 text-right">Margem se não mudar o preço</th>
                    </tr>
                  </thead>
                  <tbody>
                    {linhasMes.map((l, i) => (
                      <tr key={l.sku + i} className="border-t border-slate-100">
                        <td className="px-4 py-2 font-mono text-xs text-slate-600">{l.sku}</td>
                        <td className="px-4 py-2 text-right">{fmtR(l.venda)}</td>
                        <td className="px-4 py-2 text-right font-semibold text-emerald-700">{l.precoSugerido != null ? fmtR(l.precoSugerido) : '—'}</td>
                        <td className="px-4 py-2 text-right text-emerald-700">{fmtPct(l.margemAlvo)}</td>
                        <td className={`px-4 py-2 text-right ${l.margemSemRepasse < l.margemAlvo ? 'text-amber-600' : 'text-slate-600'}`}>{fmtPct(l.margemSemRepasse)}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}

// ─── Aba 1: Simples Nacional (hoje, 2026) ──────────────────────────────────
function SimplesNacionalTab({
  rbt12Raw, setRbt12Raw, rbt12RealRaw, setRbt12RealRaw,
  faturamentoMesRaw, setFaturamentoMesRaw, faturamentoMesRealRaw, setFaturamentoMesRealRaw,
}) {
  const rbt12 = parseBR(rbt12Raw);
  const rbt12Real = parseBR(rbt12RealRaw);
  const fatMes = parseBR(faturamentoMesRaw);
  const fatMesReal = parseBR(faturamentoMesRealRaw);
  const resultado = useMemo(() => (rbt12 > 0 ? calcDasHoje(rbt12) : null), [rbt12]);
  const resultadoReal = useMemo(() => (rbt12Real > 0 ? calcDasHoje(rbt12Real) : null), [rbt12Real]);
  const gap = resultado && resultadoReal ? resultadoReal.aliqEf - resultado.aliqEf : null;

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-slate-700 mb-1">Dados da empresa (Anexo I — Comércio)</h3>
        <p className="text-[11px] text-slate-400 mb-4">
          Preencha o RBT12 declarado (o que hoje entra no cálculo do DAS) e, se parte do faturamento não é
          declarada, o RBT12 real — pra ver o impacto de declarar 100% do faturamento, mesmo antes da reforma.
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <CurrencyInput label="RBT12 declarado (últimos 12 meses)" value={rbt12Raw} onChange={setRbt12Raw} placeholder="Ex: 480000,00" />
          <CurrencyInput label="RBT12 real (se declarasse 100%)" value={rbt12RealRaw} onChange={setRbt12RealRaw} placeholder="Ex: 800000,00" />
          <CurrencyInput label="Faturamento do mês (declarado)" value={faturamentoMesRaw} onChange={setFaturamentoMesRaw} placeholder="Ex: 40000,00" />
          <CurrencyInput label="Faturamento do mês (real)" value={faturamentoMesRealRaw} onChange={setFaturamentoMesRealRaw} placeholder="Ex: 80000,00" />
        </div>
      </div>

      {resultado && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <p className="text-xs font-semibold text-slate-500 uppercase">Faixa (declarado)</p>
            <p className="text-2xl font-black text-slate-800 mt-1">{resultado.faixaIndex + 1}ª faixa</p>
            <p className="text-[11px] text-slate-400 mt-1">até {fmtR(resultado.faixa.limite)}</p>
          </div>
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <p className="text-xs font-semibold text-slate-500 uppercase">Alíquota nominal</p>
            <p className="text-2xl font-black text-slate-800 mt-1">{fmtPct(resultado.faixa.nominal)}</p>
            <p className="text-[11px] text-slate-400 mt-1">dedução {fmtR(resultado.faixa.deducao)}</p>
          </div>
          <div className="bg-[#2F6F4E] text-white rounded-2xl border border-[#2F6F4E] p-5 shadow-sm">
            <p className="text-xs font-semibold text-emerald-100 uppercase">Alíquota efetiva hoje (declarado)</p>
            <p className="text-2xl font-black mt-1">{fmtPct(resultado.aliqEf, 4)}</p>
            <p className="text-[11px] text-emerald-100 mt-1">DAS do mês: {fmtR(fatMes * resultado.aliqEf)}</p>
          </div>
        </div>
      )}

      {resultadoReal && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <p className="text-xs font-semibold text-slate-500 uppercase">Faixa (real)</p>
            <p className="text-2xl font-black text-slate-800 mt-1">{resultadoReal.faixaIndex + 1}ª faixa</p>
            <p className="text-[11px] text-slate-400 mt-1">até {fmtR(resultadoReal.faixa.limite)}</p>
          </div>
          <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
            <p className="text-xs font-semibold text-slate-500 uppercase">Alíquota nominal (real)</p>
            <p className="text-2xl font-black text-slate-800 mt-1">{fmtPct(resultadoReal.faixa.nominal)}</p>
            <p className="text-[11px] text-slate-400 mt-1">dedução {fmtR(resultadoReal.faixa.deducao)}</p>
          </div>
          <div className="bg-amber-600 text-white rounded-2xl border border-amber-600 p-5 shadow-sm">
            <p className="text-xs font-semibold text-amber-100 uppercase">Alíquota efetiva se declarasse tudo</p>
            <p className="text-2xl font-black mt-1">{fmtPct(resultadoReal.aliqEf, 4)}</p>
            <p className="text-[11px] text-amber-100 mt-1">DAS do mês: {fmtR((fatMesReal || rbt12Real / 12) * resultadoReal.aliqEf)}</p>
          </div>
        </div>
      )}

      {gap != null && (
        <div className={`rounded-2xl border p-5 shadow-sm ${gap > 0 ? 'bg-amber-50 border-amber-200' : 'bg-emerald-50 border-emerald-200'}`}>
          {gap > 0 ? (
            <p className="text-sm text-slate-700">
              Declarando 100% do faturamento real, a alíquota efetiva subiria de <b>{fmtPct(resultado.aliqEf, 4)}</b> para{' '}
              <b>{fmtPct(resultadoReal.aliqEf, 4)}</b> (+{fmtPct(gap, 4)}){resultado.faixaIndex !== resultadoReal.faixaIndex && (
                <> — e mudaria da {resultado.faixaIndex + 1}ª pra {resultadoReal.faixaIndex + 1}ª faixa</>
              )}. Veja o impacto por produto na aba <b>Repasse</b>.
            </p>
          ) : (
            <p className="text-sm text-slate-700">Sem gap — o RBT12 real informado não é maior que o declarado.</p>
          )}
        </div>
      )}

      {resultado && (
        <p className="text-[11px] text-slate-400 italic">
          * As faixas, alíquotas nominais e deduções do Anexo I não mudam com a reforma — a LC 214/2025 só altera
          a composição interna do DAS a partir de 2027 (o que passa a ser CBS/IBS). Por isso esse cálculo vale
          tanto para hoje quanto de base para a simulação da aba "CBS e IBS".
        </p>
      )}

      {!resultado && (
        <p className="text-sm text-slate-400 italic">Informe o RBT12 declarado para calcular a faixa e a alíquota efetiva do DAS.</p>
      )}
    </div>
  );
}

// ─── Aba 2: CBS e IBS (Reforma, a partir de 2027) ──────────────────────────
function CbsIbsTab({
  rbt12Raw, setRbt12Raw, faturamentoMesRaw, setFaturamentoMesRaw,
  ano, setAno, regime, setRegime, valorVendaRaw, setValorVendaRaw,
  cbsPctRaw, setCbsPctRaw, creditosManuais, setCreditosManuais,
}) {
  const rbt12Efetivo = parseBR(rbt12Raw);
  const fatMes = parseBR(faturamentoMesRaw);
  const valorVenda = parseBR(valorVendaRaw);
  const sched = REFORM_SCHEDULE[ano];
  const cbsPct = parseBR(cbsPctRaw);

  const dentro = useMemo(() => (rbt12Efetivo > 0 ? calcDasPorDentroRate(rbt12Efetivo, ano) : null), [rbt12Efetivo, ano]);
  const fora = useMemo(() => (rbt12Efetivo > 0 ? calcDasPorForaRate(rbt12Efetivo, ano, cbsPct) : null), [rbt12Efetivo, ano, cbsPct]);

  const [formConta, setFormConta] = useState({
    descricao: '', categoria: 'mercadorias', valor: '', origemFornecedor: 'regime_normal', rbt12FornecedorRaw: '',
  });

  // O crédito de CBS/IBS não depende só da categoria da despesa — depende de
  // quem vendeu. Fornecedor Regime Normal destaca CBS/IBS integral (com o
  // limite da categoria, se houver). Fornecedor Simples Nacional só embute
  // uma fatia pequena de CBS/IBS dentro do próprio DAS dele — o crédito é
  // proporcional a essa fatia (nunca o valor cheio), e só é calculável se o
  // RBT12 do fornecedor for conhecido; sem ele, não se inventa um crédito
  // estimado. MEI não gera crédito de CBS/IBS.
  const calcCreditoConta = useCallback((conta) => {
    const { valor, categoria: categoriaId, origemFornecedor = 'regime_normal', rbt12FornecedorRaw = '' } = conta;
    const fatorCategoria = CATEGORIAS_CREDITO.find((c) => c.id === categoriaId)?.fator ?? 1;

    if (origemFornecedor === 'mei') {
      return { cbs: 0, ibs: 0, total: 0, fator: 0, motivo: 'MEI não gera crédito de CBS/IBS' };
    }
    if (origemFornecedor === 'simples_desconhecido') {
      return { cbs: 0, ibs: 0, total: 0, fator: 0, motivo: 'RBT12 do fornecedor não informado — sem crédito estimado' };
    }
    if (origemFornecedor === 'simples_conhecido') {
      const rbt12Fornecedor = parseBR(rbt12FornecedorRaw);
      if (!(rbt12Fornecedor > 0)) {
        return { cbs: 0, ibs: 0, total: 0, fator: 0, motivo: 'Informe o RBT12 do fornecedor' };
      }
      const dentroFornecedor = calcDasPorDentroRate(rbt12Fornecedor, ano);
      const cbsRateEmbutida = dentroFornecedor.detalhe.find((d) => d.nome === 'CBS')?.rate || 0;
      const ibsRateEmbutida = dentroFornecedor.detalhe.find((d) => d.nome === 'IBS')?.rate || 0;
      const cbs = valor * cbsRateEmbutida * fatorCategoria;
      const ibs = valor * ibsRateEmbutida * fatorCategoria;
      return { cbs, ibs, total: cbs + ibs, fator: fatorCategoria, proporcional: true };
    }
    // regime_normal (padrão): crédito integral, sujeito só ao limite da categoria
    const cbs = fora ? valor * fora.cbsRate * fatorCategoria : 0;
    const ibs = fora ? valor * fora.ibsRate * fatorCategoria : 0;
    return { cbs, ibs, total: cbs + ibs, fator: fatorCategoria };
  }, [fora, ano]);

  const totaisCredito = useMemo(() => creditosManuais.reduce((acc, c) => {
    const cr = calcCreditoConta(c);
    return { cbs: acc.cbs + cr.cbs, ibs: acc.ibs + cr.ibs, total: acc.total + cr.total };
  }, { cbs: 0, ibs: 0, total: 0 }), [creditosManuais, calcCreditoConta]);

  const handleAddConta = () => {
    const valor = parseBR(formConta.valor);
    if (!formConta.descricao.trim() || valor <= 0) return;
    setCreditosManuais((prev) => [...prev, {
      id: gerarId(),
      descricao: formConta.descricao.trim(),
      categoria: formConta.categoria,
      valor,
      origemFornecedor: formConta.origemFornecedor,
      rbt12FornecedorRaw: formConta.origemFornecedor === 'simples_conhecido' ? formConta.rbt12FornecedorRaw : '',
    }]);
    setFormConta((f) => ({ ...f, descricao: '', valor: '', rbt12FornecedorRaw: '' }));
  };

  const handleRemoverConta = (id) => setCreditosManuais((prev) => prev.filter((c) => c.id !== id));

  return (
    <div className="space-y-6">
      <div className="bg-sky-50 border border-sky-200 rounded-2xl p-4">
        <p className="text-xs text-sky-900">
          <b>Por que declarar tudo deixa de ser opcional a partir de 2027:</b> o CBS/IBS usa split payment — o
          imposto é segregado no momento do pagamento eletrônico da venda (cartão, Pix, boleto) e recolhido direto
          pro governo, com a nota fiscal eletrônica vinculada à transação. Fica estruturalmente mais difícil vender
          sem emitir nota ou declarar menos do que o faturamento real, diferente do Simples de hoje.
        </p>
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-slate-700 mb-1">RBT12 e faturamento</h3>
        <p className="text-[11px] text-slate-400 mb-4">
          Preenchível aqui direto — não precisa ir na aba Simples Nacional. Troque os valores à vontade pra testar
          cenários diferentes (ex: RBT12 atual, depois um RBT12 maior pra ver o efeito de declarar mais).
        </p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
          <CurrencyInput label="RBT12 (últimos 12 meses)" value={rbt12Raw} onChange={setRbt12Raw} placeholder="Ex: 480000,00" />
          <CurrencyInput label="Faturamento do mês" value={faturamentoMesRaw} onChange={setFaturamentoMesRaw} placeholder="Ex: 40000,00" />
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h3 className="text-sm font-bold text-slate-700 mb-4">Simulação da reforma</h3>
        <div className="grid grid-cols-1 sm:grid-cols-4 gap-4">
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Ano simulado</span>
            <select value={ano} onChange={(e) => setAno(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]">
              {ANOS_REFORMA.map((a) => <option key={a} value={a}>{a}</option>)}
            </select>
          </label>
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Regime</span>
            <select value={regime} onChange={(e) => setRegime(e.target.value)} className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]">
              <option value="dentro">Simples Por Dentro (sem crédito)</option>
              <option value="fora">Simples Por Fora (regime regular IBS/CBS)</option>
            </select>
          </label>
          <CurrencyInput label="Alíquota de CBS (%)" value={cbsPctRaw} onChange={setCbsPctRaw} placeholder={sched.cbs.toFixed(1)} />
          <CurrencyInput label="Valor de venda p/ simular (R$)" value={valorVendaRaw} onChange={setValorVendaRaw} placeholder="Ex: 100,00" />
        </div>
        <p className="text-[11px] text-slate-400 mt-3">
          Nominal LC 214/2025 em {ano}: CBS {sched.cbs.toFixed(1)}% · IBS {sched.ibs.toFixed(1)}% (IBS não é editável aqui).
          RBT12 usado: <b>{fmtR(rbt12Efetivo)}</b>.
        </p>
      </div>

      {rbt12Efetivo <= 0 && <p className="text-sm text-slate-400 italic">Informe o RBT12 acima para simular a reforma.</p>}

      {rbt12Efetivo > 0 && regime === 'dentro' && dentro && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
          <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">DAS reconstituído — {dentro.faixa && `${dentro.faixaIndex + 1}ª faixa`}</h4>
          <div className="grid grid-cols-2 sm:grid-cols-3 md:grid-cols-6 gap-4 mb-4">
            {dentro.detalhe.map((d) => (
              <div key={d.nome}>
                <p className="text-xs text-slate-500">{d.nome}</p>
                <p className="text-lg font-black text-slate-800">{fmtPct(d.rate, 4)}</p>
                <p className="text-[11px] text-slate-400">{fmtR(valorVenda * d.rate)}</p>
              </div>
            ))}
          </div>
          <div className="border-t border-slate-100 pt-3 flex items-baseline gap-3">
            <span className="text-3xl font-black text-[#2F6F4E]">{fmtPct(dentro.aliqEf, 4)}</span>
            <span className="text-sm text-slate-400">alíquota efetiva total (CBS + IBS embutidos)</span>
          </div>
          <p className="text-sm text-slate-600 mt-3">Sobre {fmtR(valorVenda)}: <b>{fmtR(valorVenda * dentro.aliqEf)}</b> de DAS</p>
        </div>
      )}

      {rbt12Efetivo > 0 && regime === 'fora' && fora && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
          <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">Composição — {fora.faixa && `${fora.faixaIndex + 1}ª faixa`}</h4>
          <p className="text-[10px] font-semibold text-slate-400 uppercase mb-2">Continuam no DAS residual</p>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-3">
            {fora.detalhe.map((d) => (
              <div key={d.nome}>
                <p className="text-xs text-slate-500">{d.nome}</p>
                <p className="text-lg font-black text-slate-800">{fmtPct(d.rate, 4)}</p>
                <p className="text-[11px] text-slate-400">{fmtR(valorVenda * d.rate)}</p>
              </div>
            ))}
          </div>
          <p className="text-sm text-slate-600 mb-4">
            <span className="font-black text-slate-800">{fmtPct(fora.rateResidual, 4)}</span>
            <span className="text-slate-400"> — alíquota efetiva do DAS residual (guia do DAS) · {fmtR(valorVenda * fora.rateResidual)}</span>
          </p>

          <p className="text-[10px] font-semibold text-slate-400 uppercase mb-2">Saem por fora, regime regular</p>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-4 mb-3">
            <div>
              <p className="text-xs text-slate-500">CBS ({(cbsPct || sched.cbs).toFixed(1)}%)</p>
              <p className="text-lg font-black text-amber-700">{fmtPct(fora.cbsRate)}</p>
              <p className="text-[11px] text-slate-400">{fmtR(valorVenda * fora.cbsRate)}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">IBS ({sched.ibs.toFixed(1)}%)</p>
              <p className="text-lg font-black text-amber-700">{fmtPct(fora.ibsRate)}</p>
              <p className="text-[11px] text-slate-400">{fmtR(valorVenda * fora.ibsRate)}</p>
            </div>
          </div>
          <p className="text-sm text-slate-600 mb-4">
            <span className="font-black text-amber-700">{fmtPct(fora.cbsRate + fora.ibsRate, 4)}</span>
            <span className="text-slate-400"> — alíquota efetiva de CBS + IBS (fora da guia do DAS) · {fmtR(valorVenda * (fora.cbsRate + fora.ibsRate))}</span>
          </p>

          <div className="border-t border-slate-100 pt-3 flex items-baseline gap-3">
            <span className="text-2xl font-black text-[#2F6F4E]">{fmtPct(fora.totalRate, 4)}</span>
            <span className="text-sm text-slate-400">carga total sobre a venda — DAS residual + CBS + IBS (sem créditos)</span>
          </div>
        </div>
      )}

      {rbt12Efetivo > 0 && regime === 'fora' && fora && (
        <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
          <h4 className="text-xs font-bold text-slate-500 uppercase mb-1">Créditos manuais (a apurar no mês)</h4>
          <p className="text-[11px] text-slate-400 mb-4">
            Cadastre as contas que geram crédito de CBS/IBS (compras de mercadorias, aluguel, contador...).
            Algumas categorias têm o aproveitamento limitado pela LC 214/2025 — o crédito já sai calculado com a redução aplicada.
          </p>

          <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 items-end">
            <label className="block sm:col-span-2">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Descrição da conta</span>
              <input
                type="text" value={formConta.descricao} onChange={(e) => setFormConta((f) => ({ ...f, descricao: e.target.value }))}
                placeholder="Ex: Aluguel da loja, Contador, Compra de mercadorias..."
                onKeyDown={(e) => e.key === 'Enter' && handleAddConta()}
                className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none"
              />
            </label>
            <label className="block">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Categoria</span>
              <select
                value={formConta.categoria} onChange={(e) => setFormConta((f) => ({ ...f, categoria: e.target.value }))}
                className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]"
              >
                {CATEGORIAS_CREDITO.map((c) => <option key={c.id} value={c.id}>{c.label}</option>)}
              </select>
            </label>
            <div className="flex gap-2 items-end">
              <div className="flex-1">
                <CurrencyInput label="Valor (R$)" value={formConta.valor} onChange={(v) => setFormConta((f) => ({ ...f, valor: v }))} placeholder="Ex: 3000,00" />
              </div>
              <button onClick={handleAddConta} className="h-[38px] text-sm font-semibold bg-[#2F6F4E] text-white px-4 rounded-lg hover:bg-[#265c40] transition">
                Adicionar
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 items-end mt-3">
            <label className="block sm:col-span-2">
              <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Regime do fornecedor</span>
              <select
                value={formConta.origemFornecedor} onChange={(e) => setFormConta((f) => ({ ...f, origemFornecedor: e.target.value }))}
                className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]"
              >
                <option value="regime_normal">Regime Normal (crédito integral)</option>
                <option value="simples_conhecido">Simples Nacional — RBT12 do fornecedor conhecido (crédito proporcional)</option>
                <option value="simples_desconhecido">Simples Nacional — RBT12 do fornecedor desconhecido (sem crédito)</option>
                <option value="mei">MEI (sem crédito)</option>
              </select>
            </label>
            {formConta.origemFornecedor === 'simples_conhecido' && (
              <CurrencyInput
                label="RBT12 do fornecedor (R$)" value={formConta.rbt12FornecedorRaw}
                onChange={(v) => setFormConta((f) => ({ ...f, rbt12FornecedorRaw: v }))} placeholder="Ex: 300000,00"
              />
            )}
          </div>
          <p className="text-[10px] text-slate-400 mt-2">
            Fornecedor Simples Nacional não destaca CBS/IBS separado — só uma fatia pequena já embutida no DAS dele
            gera crédito, proporcional ao RBT12 dele. Sem o RBT12 do fornecedor, não estimamos um crédito — fica zerado.
          </p>

          {creditosManuais.length > 0 && (
            <div className="mt-4 space-y-1.5">
              {creditosManuais.map((c) => {
                const cr = calcCreditoConta(c);
                const cat = CATEGORIAS_CREDITO.find((x) => x.id === c.categoria);
                return (
                  <div key={c.id} className="flex items-center justify-between gap-3 bg-slate-50 rounded-lg border border-slate-200 px-3 py-2 text-sm">
                    <div className="min-w-0">
                      <p className="text-slate-700 font-semibold truncate">{c.descricao}</p>
                      <p className="text-[11px] text-slate-400">
                        {cat?.label} · base {fmtR(c.valor)}
                        {cr.fator < 1 && cr.total > 0 && <span className="text-amber-600 font-semibold"> · crédito {fmtPct(cr.fator, 0)}</span>}
                        {cr.proporcional && <span className="text-sky-600 font-semibold"> · proporcional (fornecedor Simples)</span>}
                        {cr.motivo && <span className="text-red-500 font-semibold"> · {cr.motivo}</span>}
                      </p>
                    </div>
                    <div className="flex items-center gap-3 shrink-0">
                      <span className="text-xs text-slate-500">CBS {fmtR(cr.cbs)}</span>
                      <span className="text-xs text-slate-500">IBS {fmtR(cr.ibs)}</span>
                      <span className="font-semibold text-emerald-700">{fmtR(cr.total)}</span>
                      <button onClick={() => handleRemoverConta(c.id)} className="text-xs text-slate-300 hover:text-red-500">remover</button>
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 mt-4 border-t border-slate-100 pt-4">
            <div>
              <p className="text-xs text-slate-500">CBS a pagar no mês (líquido)</p>
              <p className="text-xl font-black text-slate-800">{fmtR(Math.max(0, fatMes * fora.cbsRate - totaisCredito.cbs))}</p>
              <p className="text-[11px] text-slate-400">bruto {fmtR(fatMes * fora.cbsRate)} − crédito {fmtR(totaisCredito.cbs)}</p>
            </div>
            <div>
              <p className="text-xs text-slate-500">IBS a pagar no mês (líquido)</p>
              <p className="text-xl font-black text-slate-800">{fmtR(Math.max(0, fatMes * fora.ibsRate - totaisCredito.ibs))}</p>
              <p className="text-[11px] text-slate-400">bruto {fmtR(fatMes * fora.ibsRate)} − crédito {fmtR(totaisCredito.ibs)}</p>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

// ─── Aba: Precificação ──────────────────────────────────────────────────────
function cenario(venda, custo, margemHoje, rate) {
  const denom = 1 - margemHoje - rate;
  const precoSugerido = denom > 0 ? custo / denom : null;
  const aumentoValor = precoSugerido != null ? precoSugerido - venda : null;
  const aumentoPct = precoSugerido != null && venda > 0 ? aumentoValor / venda : null;
  return { precoSugerido, aumentoValor, aumentoPct };
}

function AumentoCell({ c }) {
  if (c.aumentoPct == null) return <span className="text-red-500">inviável</span>;
  return (
    <span className={c.aumentoPct > 0 ? 'text-amber-600 font-semibold' : 'text-slate-500'}>
      {c.aumentoValor >= 0 ? '+' : ''}{fmtR(c.aumentoValor)} ({fmtPct(c.aumentoPct)})
    </span>
  );
}

// ─── Aba: Produtos (cadastro + custos, em cards) ───────────────────────────
function ProdutoModal({ produto, onClose, onAtualizarCustoBase, onAdicionarCustoExtra, onRemoverCustoExtra }) {
  const [novoCustoNome, setNovoCustoNome] = useState('');
  const [novoCustoValor, setNovoCustoValor] = useState('');

  const handleAdicionar = () => {
    onAdicionarCustoExtra(produto.id, novoCustoNome, novoCustoValor);
    setNovoCustoNome('');
    setNovoCustoValor('');
  };

  return (
    <div className="fixed inset-0 bg-slate-900/40 flex items-center justify-center p-4 z-50" onClick={onClose}>
      <div className="bg-white rounded-2xl shadow-xl w-full max-w-lg p-6" onClick={(e) => e.stopPropagation()}>
        <div className="flex items-start justify-between mb-4">
          <div>
            <p className="font-mono text-sm font-bold text-slate-800">{produto.sku}</p>
            <p className="text-[11px] text-slate-400 mt-0.5">Preço de venda: {fmtR(produto.venda)}</p>
          </div>
          <button onClick={onClose} className="text-slate-400 hover:text-slate-600 text-lg leading-none">✕</button>
        </div>

        <label className="block">
          <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Preço de compra (base)</span>
          <input
            type="text" inputMode="decimal" defaultValue={String(produto.custoBase).replace('.', ',')}
            onBlur={(e) => onAtualizarCustoBase(produto.id, e.target.value)}
            className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none"
          />
        </label>

        {produto.custosExtras?.length > 0 && (
          <div className="mt-4">
            <p className="text-xs font-semibold text-slate-500 uppercase mb-2">Custos extras</p>
            <div className="space-y-1.5">
              {produto.custosExtras.map((c) => (
                <div key={c.id} className="flex items-center justify-between bg-slate-50 rounded-lg border border-slate-200 px-3 py-1.5 text-sm">
                  <span className="text-slate-600">{c.nome}</span>
                  <div className="flex items-center gap-3">
                    <span className="font-semibold text-slate-700">{fmtR(c.valor)}</span>
                    <button onClick={() => onRemoverCustoExtra(produto.id, c.id)} className="text-xs text-slate-300 hover:text-red-500">remover</button>
                  </div>
                </div>
              ))}
            </div>
          </div>
        )}

        <div className="mt-4 flex flex-wrap items-end gap-3">
          <label className="block flex-1 min-w-[140px]">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">Novo custo</span>
            <input
              type="text" value={novoCustoNome} onChange={(e) => setNovoCustoNome(e.target.value)} placeholder="Ex: Frete, embalagem, taxa..."
              className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none"
            />
          </label>
          <div className="w-28">
            <CurrencyInput label="Valor (R$)" value={novoCustoValor} onChange={setNovoCustoValor} placeholder="Ex: 3,50" />
          </div>
          <button onClick={handleAdicionar} className="h-[38px] text-sm font-semibold bg-[#2F6F4E] text-white px-4 rounded-lg hover:bg-[#265c40] transition">
            Adicionar
          </button>
        </div>

        <div className="mt-5 border-t border-slate-100 pt-4 flex items-baseline justify-between">
          <span className="text-xs font-semibold text-slate-500 uppercase">Custo total</span>
          <span className="text-xl font-black text-[#2F6F4E]">{fmtR(custoTotal({ custo: produto.custoBase, custosExtras: produto.custosExtras }))}</span>
        </div>
      </div>
    </div>
  );
}

function ProdutosTab({ produtos, setProdutos }) {
  const fileInputRef = useRef(null);
  const [novoSku, setNovoSku] = useState('');
  const [novoCusto, setNovoCusto] = useState('');
  const [novoVenda, setNovoVenda] = useState('');
  const [abertoId, setAbertoId] = useState(null);

  const handleFile = useCallback((file) => {
    const reader = new FileReader();
    reader.onload = async (e) => {
      try {
        const parsed = parseProdutos(e.target.result);
        const salvos = await substituirTodosProdutos(parsed);
        setProdutos(salvos);
      } catch (err) {
        alert('Não foi possível ler essa planilha: ' + err.message);
      }
    };
    reader.readAsArrayBuffer(file);
  }, [setProdutos]);

  const adicionarProduto = useCallback(async () => {
    const custo = parseBR(novoCusto);
    const venda = parseBR(novoVenda);
    if (!novoSku.trim() || venda <= 0) return;
    const novo = await criarProduto({ sku: novoSku.trim(), custo, venda });
    setProdutos((prev) => [...prev, novo]);
    setNovoSku('');
    setNovoCusto('');
    setNovoVenda('');
  }, [novoSku, novoCusto, novoVenda, setProdutos]);

  const removerProduto = useCallback((id) => {
    setProdutos((prev) => prev.filter((p) => p.id !== id));
    excluirProduto(id);
  }, [setProdutos]);

  const handleAtualizarCustoBase = useCallback((id, valor) => {
    const custo = parseBR(valor);
    setProdutos((prev) => prev.map((p) => (p.id === id ? { ...p, custo } : p)));
    atualizarCustoBase(id, custo);
  }, [setProdutos]);

  const handleAdicionarCustoExtra = useCallback(async (id, nome, valorRaw) => {
    const valor = parseBR(valorRaw);
    if (!nome.trim() || valor <= 0) return;
    const novo = await criarCustoExtra(id, nome.trim(), valor);
    setProdutos((prev) => prev.map((p) => (
      p.id === id ? { ...p, custosExtras: [...(p.custosExtras || []), novo] } : p
    )));
  }, [setProdutos]);

  const removerCustoExtra = useCallback((produtoId, custoId) => {
    setProdutos((prev) => prev.map((p) => (
      p.id === produtoId ? { ...p, custosExtras: (p.custosExtras || []).filter((c) => c.id !== custoId) } : p
    )));
    excluirCustoExtra(produtoId, custoId);
  }, [setProdutos]);

  const produtoAberto = produtos.find((p) => p.id === abertoId);

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-wrap items-center justify-between gap-3">
        <h3 className="text-sm font-bold text-slate-700">Produtos ({produtos.length})</h3>
        <div className="flex items-center gap-3">
          <button
            onClick={() => fileInputRef.current?.click()}
            className="flex items-center gap-2 text-sm font-semibold bg-[#2F6F4E] text-white px-4 py-2 rounded-lg hover:bg-[#265c40] transition"
          >
            <Upload size={16} /> Importar planilha
          </button>
          <input ref={fileInputRef} type="file" accept=".xlsx,.xls" className="hidden" onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])} />
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm">
        <h4 className="text-xs font-bold text-slate-500 uppercase mb-3">Adicionar produto manualmente</h4>
        <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 items-end">
          <label className="block">
            <span className="text-xs font-semibold text-slate-500 uppercase tracking-wide">SKU / Produto</span>
            <input type="text" value={novoSku} onChange={(e) => setNovoSku(e.target.value)} placeholder="Ex: KIT2NOVO-M"
              className="mt-1 w-full rounded-lg border border-slate-300 px-3 py-2 text-sm focus:border-[#2F6F4E] focus:ring-1 focus:ring-[#2F6F4E] outline-none" />
          </label>
          <CurrencyInput label="Preço de compra (custo)" value={novoCusto} onChange={setNovoCusto} placeholder="Ex: 15,00" />
          <CurrencyInput label="Preço de venda" value={novoVenda} onChange={setNovoVenda} placeholder="Ex: 35,00" />
          <button onClick={adicionarProduto} className="h-[38px] text-sm font-semibold bg-[#2F6F4E] text-white px-4 py-2 rounded-lg hover:bg-[#265c40] transition">
            Adicionar
          </button>
        </div>
      </div>

      <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
        {produtos.length === 0 ? (
          <p className="text-sm text-slate-400 italic text-center py-8">Nenhum produto carregado. Importe uma planilha ou adicione manualmente.</p>
        ) : (
          <div className="divide-y divide-slate-100">
            {produtos.map((p) => (
              <div
                key={p.id}
                onClick={() => setAbertoId(p.id)}
                className="flex items-center justify-between gap-4 px-5 py-3 hover:bg-slate-50 transition cursor-pointer group"
              >
                <div className="min-w-0 flex items-center gap-2">
                  <p className="font-mono text-xs font-bold text-slate-700 truncate">{p.sku}</p>
                  {p.custosExtras?.length > 0 && (
                    <span className="shrink-0 text-[10px] font-sans bg-slate-100 text-slate-500 px-2 py-0.5 rounded-full">
                      +{p.custosExtras.length} custo{p.custosExtras.length > 1 ? 's' : ''} extra{p.custosExtras.length > 1 ? 's' : ''}
                    </span>
                  )}
                </div>
                <div className="flex items-center gap-6 shrink-0">
                  <div className="text-right">
                    <span className="text-[10px] text-slate-400 block">Venda</span>
                    <span className="text-sm font-semibold text-slate-700">{fmtR(p.venda)}</span>
                  </div>
                  <div className="text-right">
                    <span className="text-[10px] text-slate-400 block">Custo total</span>
                    <span className="text-sm font-semibold text-slate-700">{fmtR(custoTotal(p))}</span>
                  </div>
                  <button
                    onClick={(e) => { e.stopPropagation(); removerProduto(p.id); }}
                    className="text-xs text-slate-300 hover:text-red-500 opacity-0 group-hover:opacity-100 transition"
                  >
                    remover
                  </button>
                </div>
              </div>
            ))}
          </div>
        )}
      </div>

      {produtoAberto && (
        <ProdutoModal
          produto={{ ...produtoAberto, custoBase: produtoAberto.custo }}
          onClose={() => setAbertoId(null)}
          onAtualizarCustoBase={handleAtualizarCustoBase}
          onAdicionarCustoExtra={handleAdicionarCustoExtra}
          onRemoverCustoExtra={removerCustoExtra}
        />
      )}
    </div>
  );
}

// ─── Aba: Precificação (comparação Hoje x Se declarar tudo x Por Dentro x Por Fora) ──
function PrecificacaoTab({ rbt12Raw, rbt12RealRaw, ano, setAno, cbsPctRaw, produtos }) {
  const rbt12Declarado = parseBR(rbt12Raw);
  const rbt12Real = parseBR(rbt12RealRaw);
  const rbt12Efetivo = rbt12Real > 0 ? rbt12Real : rbt12Declarado;
  const cbsPct = parseBR(cbsPctRaw);

  const hoje = useMemo(() => (rbt12Declarado > 0 ? calcDasHoje(rbt12Declarado) : null), [rbt12Declarado]);
  const hojeFormalizado = useMemo(() => (rbt12Real > 0 ? calcDasHoje(rbt12Real) : null), [rbt12Real]);
  const dentroInfo = useMemo(() => (rbt12Efetivo > 0 ? calcDasPorDentroRate(rbt12Efetivo, ano) : null), [rbt12Efetivo, ano]);
  const foraInfo = useMemo(() => (rbt12Efetivo > 0 ? calcDasPorForaRate(rbt12Efetivo, ano, cbsPct) : null), [rbt12Efetivo, ano, cbsPct]);

  const linhas = useMemo(() => {
    if (!hoje || !dentroInfo || !foraInfo) return [];
    return produtos.map((p) => {
      const custo = custoTotal(p);
      const dasHoje = p.venda * hoje.aliqEf;
      const lucroHoje = p.venda - custo - dasHoje;
      const margemHoje = p.venda > 0 ? lucroHoje / p.venda : 0;
      return {
        ...p,
        custo,
        margemHoje,
        formalizado: hojeFormalizado ? cenario(p.venda, custo, margemHoje, hojeFormalizado.aliqEf) : null,
        dentro: cenario(p.venda, custo, margemHoje, dentroInfo.aliqEf),
        fora: cenario(p.venda, custo, margemHoje, foraInfo.totalRate),
      };
    });
  }, [produtos, hoje, hojeFormalizado, dentroInfo, foraInfo]);

  return (
    <div className="space-y-6">
      <div className="bg-white rounded-2xl border border-slate-200 p-5 shadow-sm flex flex-wrap items-center justify-between gap-3">
        <div>
          <h3 className="text-sm font-bold text-slate-700">Produtos ({produtos.length})</h3>
          <p className="text-[11px] text-slate-400 mt-1">Comparando DAS de hoje x se declarasse tudo x Simples Por Dentro x Simples Por Fora em</p>
        </div>
        <select value={ano} onChange={(e) => setAno(e.target.value)} className="rounded-lg border border-slate-300 px-3 py-2 text-sm outline-none focus:border-[#2F6F4E]">
          {ANOS_REFORMA.map((a) => <option key={a} value={a}>{a}</option>)}
        </select>
      </div>

      {!hoje && <p className="text-sm text-slate-400 italic">Informe o RBT12 na aba Simples Nacional para calcular a precificação.</p>}

      {hoje && !hojeFormalizado && (
        <p className="text-[11px] text-slate-400 italic">
          Informe também o RBT12 real na aba Simples Nacional pra ver a coluna "Se declarar tudo" — o reajuste
          necessário só por causa da formalização, antes mesmo da reforma.
        </p>
      )}

      {hoje && (
        <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-x-auto">
          <table className="w-full text-sm">
            <thead>
              <tr className="text-left text-[11px] uppercase text-slate-500">
                <th className="px-4 py-2 bg-slate-50" rowSpan={2}>SKU</th>
                <th className="px-4 py-2 text-right bg-slate-50" rowSpan={2}>Custo</th>
                <th className="px-4 py-2 text-right bg-slate-50" colSpan={2}>Hoje</th>
                <th className="px-4 py-2 text-right bg-sky-50" colSpan={2}>Se declarar tudo (hoje)</th>
                <th className="px-4 py-2 text-right bg-emerald-50" colSpan={2}>Por Dentro ({ano})</th>
                <th className="px-4 py-2 text-right bg-amber-50" colSpan={2}>Por Fora ({ano})</th>
              </tr>
              <tr className="text-left text-[11px] uppercase text-slate-500">
                <th className="px-4 py-1 text-right bg-slate-50">Preço atual</th>
                <th className="px-4 py-1 text-right bg-slate-50">Margem líquida</th>
                <th className="px-4 py-1 text-right bg-sky-50">Preço sugerido</th>
                <th className="px-4 py-1 text-right bg-sky-50">Aumento</th>
                <th className="px-4 py-1 text-right bg-emerald-50">Preço sugerido</th>
                <th className="px-4 py-1 text-right bg-emerald-50">Aumento</th>
                <th className="px-4 py-1 text-right bg-amber-50">Preço sugerido</th>
                <th className="px-4 py-1 text-right bg-amber-50">Aumento</th>
              </tr>
            </thead>
            <tbody>
              {linhas.map((l) => (
                <tr key={l.id} className="border-t border-slate-100">
                  <td className="px-4 py-2 font-mono text-xs text-slate-600">
                    {l.sku}
                    {l.custosExtras?.length > 0 && (
                      <span className="ml-1.5 text-[9px] font-sans bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded-full">+{l.custosExtras.length}</span>
                    )}
                  </td>
                  <td className="px-4 py-2 text-right">{fmtR(l.custo)}</td>
                  <td className="px-4 py-2 text-right">{fmtR(l.venda)}</td>
                  <td className={`px-4 py-2 text-right font-semibold ${l.margemHoje < 0 ? 'text-red-500' : 'text-slate-700'}`}>{fmtPct(l.margemHoje)}</td>
                  <td className="px-4 py-2 text-right font-semibold text-sky-700">{l.formalizado?.precoSugerido != null ? fmtR(l.formalizado.precoSugerido) : '—'}</td>
                  <td className="px-4 py-2 text-right">{l.formalizado ? <AumentoCell c={l.formalizado} /> : <span className="text-slate-300">—</span>}</td>
                  <td className="px-4 py-2 text-right font-semibold text-emerald-700">{l.dentro.precoSugerido != null ? fmtR(l.dentro.precoSugerido) : '—'}</td>
                  <td className="px-4 py-2 text-right"><AumentoCell c={l.dentro} /></td>
                  <td className="px-4 py-2 text-right font-semibold text-amber-700">{l.fora.precoSugerido != null ? fmtR(l.fora.precoSugerido) : '—'}</td>
                  <td className="px-4 py-2 text-right"><AumentoCell c={l.fora} /></td>
                </tr>
              ))}
              {linhas.length === 0 && (
                <tr><td colSpan={10} className="px-4 py-8 text-center text-slate-400 italic">Nenhum produto carregado — vá na aba Produtos.</td></tr>
              )}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

export default function App() {
  const [tab, setTab] = useState('simples');
  const [rbt12Raw, setRbt12Raw] = useState('');
  const [rbt12RealRaw, setRbt12RealRaw] = useState('');
  const [faturamentoMesRaw, setFaturamentoMesRaw] = useState('');
  const [faturamentoMesRealRaw, setFaturamentoMesRealRaw] = useState('');
  const [ano, setAno] = useState('2027');
  const [regime, setRegime] = useState('fora');
  const [valorVendaRaw, setValorVendaRaw] = useState('100,00');
  const [cbsPctRaw, setCbsPctRaw] = useState(String(REFORM_SCHEDULE['2027'].cbs));
  const [creditosManuais, setCreditosManuais] = useState([]);
  const [aliquotaUsadaRaw, setAliquotaUsadaRaw] = useState('');
  const [mesesRaw, setMesesRaw] = useState('5');
  const [produtos, setProdutos] = useState([]);

  // Ao trocar o ano simulado, repõe a alíquota de CBS para a nominal oficial
  // daquele ano — o usuário pode sobrescrever de novo em seguida.
  useEffect(() => {
    setCbsPctRaw(String(REFORM_SCHEDULE[ano].cbs));
  }, [ano]);

  // Carrega os produtos: Supabase primeiro (se configurado — VITE_SUPABASE_URL
  // / VITE_SUPABASE_ANON_KEY no .env), senão o que estiver salvo no navegador;
  // sem nada em nenhum dos dois (primeiríssima vez), semeia a partir da
  // planilha-base — e já grava essa semente no banco, se configurado.
  useEffect(() => {
    carregarOuSemear(() =>
      fetch(produtosBaseUrl).then((r) => r.arrayBuffer()).then(parseProdutos)
    ).then(setProdutos).catch(() => {});
  }, []);

  // Cópia local sempre atualizada — cache/fallback caso o Supabase esteja
  // fora do ar ou não configurado; as gravações "de verdade" acontecem por
  // operação (criar/editar/excluir), não aqui.
  useEffect(() => {
    if (produtos.length) salvarSnapshotLocal(produtos);
  }, [produtos]);

  const tabs = [
    { id: 'simples', label: 'Simples Nacional', icon: Calculator },
    { id: 'reforma', label: 'CBS e IBS', icon: TrendingUp },
    { id: 'produtos', label: 'Produtos', icon: Package },
    { id: 'precificacao', label: 'Precificação', icon: FileSpreadsheet },
    { id: 'repasse', label: 'Repasse', icon: LineChart },
  ];

  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-white border-b border-slate-200 px-6 py-4">
        <h1 className="text-lg font-black text-slate-800">Precificador — Simples x Reforma Tributária</h1>
        <p className="text-xs text-slate-400 mt-0.5">Comércio · Anexo I · entrada manual</p>
      </header>

      <nav className="bg-white border-b border-slate-200 px-6">
        <div className="flex gap-1">
          {tabs.map(({ id, label, icon: Icon }) => (
            <button
              key={id}
              onClick={() => setTab(id)}
              className={`flex items-center gap-2 px-4 py-3 text-sm font-semibold border-b-2 transition ${
                tab === id ? 'border-[#2F6F4E] text-[#2F6F4E]' : 'border-transparent text-slate-400 hover:text-slate-600'
              }`}
            >
              <Icon size={16} /> {label}
            </button>
          ))}
        </div>
      </nav>

      <main className="max-w-6xl mx-auto px-6 py-6">
        {tab === 'simples' && (
          <SimplesNacionalTab
            rbt12Raw={rbt12Raw}
            setRbt12Raw={setRbt12Raw}
            rbt12RealRaw={rbt12RealRaw}
            setRbt12RealRaw={setRbt12RealRaw}
            faturamentoMesRaw={faturamentoMesRaw}
            setFaturamentoMesRaw={setFaturamentoMesRaw}
            faturamentoMesRealRaw={faturamentoMesRealRaw}
            setFaturamentoMesRealRaw={setFaturamentoMesRealRaw}
          />
        )}
        {tab === 'reforma' && (
          <CbsIbsTab
            rbt12Raw={rbt12Raw}
            setRbt12Raw={setRbt12Raw}
            faturamentoMesRaw={faturamentoMesRaw}
            setFaturamentoMesRaw={setFaturamentoMesRaw}
            ano={ano}
            setAno={setAno}
            regime={regime}
            setRegime={setRegime}
            valorVendaRaw={valorVendaRaw}
            setValorVendaRaw={setValorVendaRaw}
            cbsPctRaw={cbsPctRaw}
            setCbsPctRaw={setCbsPctRaw}
            creditosManuais={creditosManuais}
            setCreditosManuais={setCreditosManuais}
          />
        )}
        {tab === 'produtos' && (
          <ProdutosTab produtos={produtos} setProdutos={setProdutos} />
        )}
        {tab === 'precificacao' && (
          <PrecificacaoTab
            rbt12Raw={rbt12Raw}
            rbt12RealRaw={rbt12RealRaw}
            ano={ano}
            setAno={setAno}
            cbsPctRaw={cbsPctRaw}
            produtos={produtos}
          />
        )}
        {tab === 'repasse' && (
          <RepasseTab
            rbt12Raw={rbt12Raw}
            rbt12RealRaw={rbt12RealRaw}
            produtos={produtos}
            aliquotaUsadaRaw={aliquotaUsadaRaw}
            setAliquotaUsadaRaw={setAliquotaUsadaRaw}
            mesesRaw={mesesRaw}
            setMesesRaw={setMesesRaw}
          />
        )}
      </main>
    </div>
  );
}
