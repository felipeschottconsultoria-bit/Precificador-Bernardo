// Persistência dos produtos: usa o Firestore quando as chaves estão
// configuradas (.env — VITE_FIREBASE_*); sem isso, cai automaticamente para
// o localStorage do navegador (mesmo comportamento de antes) — o app nunca
// fica sem salvar nada.
//
// Modelagem: cada produto é um documento na coleção `produtos`, com os
// custos extras embutidos como um array dentro do próprio documento (em vez
// de uma subcoleção) — mais simples de ler (1 leitura por produto, sem
// precisar buscar subcoleções) e suficiente pro volume desse catálogo.
import { db, firebaseConfigurado } from './firebaseClient';
import {
  collection, doc, getDocs, addDoc, updateDoc, deleteDoc, writeBatch, arrayUnion, arrayRemove,
} from 'firebase/firestore';

const LS_KEY = 'precificador-reforma:produtos';
const COLECAO = 'produtos';

export const gerarId = () =>
  (typeof crypto !== 'undefined' && crypto.randomUUID)
    ? crypto.randomUUID()
    : `${Date.now()}-${Math.random().toString(36).slice(2, 9)}`;

// custo de compra + soma de todos os custos extras adicionados no produto.
export const custoTotal = (p) => (p.custo || 0) + (p.custosExtras || []).reduce((acc, c) => acc + (c.valor || 0), 0);

const carregarProdutosLocal = () => {
  try {
    const raw = localStorage.getItem(LS_KEY);
    const parsed = raw ? JSON.parse(raw) : null;
    return Array.isArray(parsed) && parsed.length ? parsed : null;
  } catch {
    return null;
  }
};

export const salvarSnapshotLocal = (produtos) => {
  try {
    localStorage.setItem(LS_KEY, JSON.stringify(produtos));
  } catch {
    // localStorage indisponível (modo privado, quota cheia) — segue sem salvar.
  }
};

const docParaProduto = (d) => {
  const data = d.data();
  return { id: d.id, sku: data.sku, venda: data.venda, custo: data.custo, custosExtras: data.custosExtras || [] };
};

// Ao iniciar: Firestore primeiro (fonte de verdade, se configurado); sem
// documentos lá ainda (banco novo) ou sem Firebase configurado, cai pro
// localStorage; sem nada em nenhum dos dois, devolve null (chamador semeia
// a partir da planilha-base).
export const carregarProdutos = async () => {
  if (firebaseConfigurado) {
    const snap = await getDocs(collection(db, COLECAO));
    if (!snap.empty) return snap.docs.map(docParaProduto);
  }
  return carregarProdutosLocal();
};

// Semeia o banco (Firestore, se configurado) com a lista inicial — usado só
// na primeira carga, quando não há nada salvo em lugar nenhum ainda.
export const semearProdutos = async (produtos) => {
  if (!firebaseConfigurado) return produtos;
  const batch = writeBatch(db);
  const refs = produtos.map((p) => {
    const ref = doc(collection(db, COLECAO));
    batch.set(ref, { sku: p.sku, venda: p.venda, custo: p.custo, custosExtras: [] });
    return ref;
  });
  await batch.commit();
  return refs.map((ref, i) => ({ id: ref.id, sku: produtos[i].sku, venda: produtos[i].venda, custo: produtos[i].custo, custosExtras: [] }));
};

// Carga inicial memoizada num promise em nível de módulo — o React
// StrictMode (dev) roda o efeito de montagem duas vezes de propósito; sem
// isso, as duas chamadas assíncronas corririam em paralelo e, achando o
// banco vazio nas duas, semeariam a planilha-base duas vezes (produtos
// duplicados). O módulo só existe uma vez por carregamento de página, então
// esse cache garante uma única execução mesmo com o duplo-efeito do
// StrictMode ou cliques repetidos.
let cargaInicialPromise = null;
export const carregarOuSemear = (buscarSeed) => {
  if (!cargaInicialPromise) {
    cargaInicialPromise = (async () => {
      const salvos = await carregarProdutos();
      if (salvos) return salvos;
      const seed = await buscarSeed();
      return semearProdutos(seed);
    })();
  }
  return cargaInicialPromise;
};

export const criarProduto = async ({ sku, custo, venda }) => {
  if (firebaseConfigurado) {
    const ref = await addDoc(collection(db, COLECAO), { sku, venda, custo, custosExtras: [] });
    return { id: ref.id, sku, venda, custo, custosExtras: [] };
  }
  return { id: gerarId(), sku, venda, custo, custosExtras: [] };
};

export const excluirProduto = async (id) => {
  if (firebaseConfigurado) await deleteDoc(doc(db, COLECAO, id));
};

export const atualizarCustoBase = async (id, custo) => {
  if (firebaseConfigurado) await updateDoc(doc(db, COLECAO, id), { custo });
};

export const criarCustoExtra = async (produtoId, nome, valor) => {
  const novo = { id: gerarId(), nome, valor };
  if (firebaseConfigurado) {
    await updateDoc(doc(db, COLECAO, produtoId), { custosExtras: arrayUnion(novo) });
  }
  return novo;
};

export const excluirCustoExtra = async (produtoId, custoId) => {
  if (!firebaseConfigurado) return;
  // arrayRemove precisa do objeto exato — busca o produto atual pra achar o
  // item certo antes de remover (não dá pra remover só pelo id direto).
  const snap = await getDocs(collection(db, COLECAO));
  const produtoDoc = snap.docs.find((d) => d.id === produtoId);
  if (!produtoDoc) return;
  const alvo = (produtoDoc.data().custosExtras || []).find((c) => c.id === custoId);
  if (!alvo) return;
  await updateDoc(doc(db, COLECAO, produtoId), { custosExtras: arrayRemove(alvo) });
};

// Importar planilha substitui a lista inteira — apaga tudo e insere de novo.
export const substituirTodosProdutos = async (produtos) => {
  if (!firebaseConfigurado) return produtos.map((p) => ({ ...p, id: gerarId(), custosExtras: [] }));
  const existentes = await getDocs(collection(db, COLECAO));
  const batchDelete = writeBatch(db);
  existentes.docs.forEach((d) => batchDelete.delete(d.ref));
  await batchDelete.commit();
  return semearProdutos(produtos);
};
