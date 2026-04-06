// Planilha de Oportunidade - Utility Functions

import type { OpportunityRecord } from './types';
import { COLUMN_MAP, SHEET_NAMES } from './constants';

export function parseDate(value: unknown): Date | null {
  if (!value) return null;

  if (typeof value === 'number') {
    const excelEpoch = new Date(1899, 11, 30);
    return new Date(excelEpoch.getTime() + value * 24 * 60 * 60 * 1000);
  }

  const dateStr = String(value).trim();

  const match1 = dateStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match1) {
    const [, day, month, year] = match1;
    return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
  }

  const match2 = dateStr.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (match2) {
    const [, year, month, day] = match2;
    return new Date(parseInt(year), parseInt(month) - 1, parseInt(day));
  }

  const date = new Date(dateStr);
  return isNaN(date.getTime()) ? null : date;
}

export function parseNumber(value: unknown): number {
  if (value === null || value === undefined) return 0;
  if (typeof value === 'number') return value;
  if (typeof value === 'string') {
    const trimmed = value.trim();
    if (trimmed.startsWith('#') || trimmed === '' || trimmed === '-') return 0;
    const normalized = trimmed.replace(',', '.').trim();
    const num = parseFloat(normalized);
    return isNaN(num) ? 0 : num;
  }
  return 0;
}

export function parseCriticidade(value: string): string {
  const raw = value.toLowerCase().trim();
  if (raw === 'alta' || raw === 'alto') return 'Alta';
  if (raw === 'média' || raw === 'media' || raw === 'medio' || raw === 'médio') return 'Média';
  return 'Baixa';
}

export function determineStatus(
  notaFiscal: string,
  dataVenda: Date | null,
  pedidoCompra: string,
  ordemManutencao: string,
  quantidadeFaturada: number = 0,
  quantidadePedida: number = 0,
  qty: number = 0
): OpportunityRecord['status'] {
  const qtdPedida = quantidadePedida || qty;
  const hasFaturamento = quantidadeFaturada > 0;

  if (hasFaturamento && qtdPedida > 0 && quantidadeFaturada < qtdPedida) return 'faturado_parcial';
  if (hasFaturamento && (qtdPedida === 0 || quantidadeFaturada >= qtdPedida)) return 'vendido';
  if (notaFiscal || dataVenda) return 'vendido';
  if (pedidoCompra) return 'com_po_sem_faturamento';
  if (ordemManutencao) return 'com_om';
  return 'sem_om';
}

export function getStockInfo(record: OpportunityRecord) {
  const qtdPedida = record.quantidadePedida || record.qty;
  const isLundin = record.origemAba?.toLowerCase().includes('lundin');
  const importacaoQtd = record.importacao || 0;

  let estoqueDisponivel = 0;
  let localEstoque = '';
  let precisaTransferencia = false;

  if (isLundin) {
    if (record.lic >= qtdPedida) {
      estoqueDisponivel = record.lic;
      localEstoque = 'LIC';
    } else if (record.betim >= qtdPedida) {
      estoqueDisponivel = record.betim;
      localEstoque = 'Betim';
      precisaTransferencia = true;
    } else if ((record.lic + record.betim) >= qtdPedida) {
      estoqueDisponivel = record.lic + record.betim;
      localEstoque = 'LIC+Betim';
      precisaTransferencia = true;
    } else if (importacaoQtd >= qtdPedida) {
      estoqueDisponivel = importacaoQtd;
      localEstoque = 'Importação';
    } else {
      estoqueDisponivel = record.lic + record.betim + importacaoQtd;
      localEstoque = 'Insuficiente';
    }
  } else {
    if (record.betim >= qtdPedida) {
      estoqueDisponivel = record.betim;
      localEstoque = 'Betim';
    } else if (importacaoQtd >= qtdPedida) {
      estoqueDisponivel = importacaoQtd;
      localEstoque = 'Importação';
    } else if ((record.betim + importacaoQtd) >= qtdPedida) {
      estoqueDisponivel = record.betim + importacaoQtd;
      localEstoque = 'Betim+Imp';
    } else {
      estoqueDisponivel = record.betim + importacaoQtd;
      localEstoque = 'Insuficiente';
    }
  }

  const saldo = qtdPedida - estoqueDisponivel;
  const coberturaTotal = estoqueDisponivel >= qtdPedida;

  let bgColor = '#dc3545';
  let textColor = 'white';

  if (coberturaTotal && !precisaTransferencia) {
    bgColor = '#28a745';
  } else if (coberturaTotal && precisaTransferencia) {
    bgColor = '#ffc107';
    textColor = '#333';
  } else if (saldo > 0 && estoqueDisponivel > 0) {
    bgColor = '#fd7e14';
  }

  return { estoqueDisponivel, localEstoque, precisaTransferencia, saldo, coberturaTotal, bgColor, textColor };
}

export function parseExcelFile(fileData: ArrayBuffer): OpportunityRecord[] {
  // eslint-disable-next-line @typescript-eslint/no-require-imports
  const XLSX = require('xlsx');
  const workbook = XLSX.read(fileData);
  const allParsedData: OpportunityRecord[] = [];
  let globalId = 1;

  for (const sheetName of SHEET_NAMES) {
    const actualSheetName = workbook.SheetNames.find(
      (name: string) => name.toLowerCase() === sheetName.toLowerCase()
    );

    if (!actualSheetName) {
      console.warn(`Aba "${sheetName}" não encontrada`);
      continue;
    }

    const worksheet = workbook.Sheets[actualSheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as unknown[][];

    if (jsonData.length < 2) continue;

    const headers = jsonData[0] as string[];
    const columnIndices: Record<string, number> = {};

    headers.forEach((header, index) => {
      const normalizedHeader = header?.toString().trim();
      const mappedKey = COLUMN_MAP[normalizedHeader];
      if (mappedKey) {
        columnIndices[mappedKey] = index;
      }
    });

    for (let i = 1; i < jsonData.length; i++) {
      const row = jsonData[i] as unknown[];
      if (!row || row.length === 0) continue;

      const getValue = (key: string): unknown => {
        const idx = columnIndices[key];
        return idx !== undefined ? row[idx] : undefined;
      };

      const empresa = String(getValue('empresa') || '').trim();
      const cliente = String(getValue('cliente') || '').trim();
      const qty = parseNumber(getValue('qty'));
      const lic = parseNumber(getValue('lic'));
      const betim = parseNumber(getValue('betim'));
      const criticidade = parseCriticidade(String(getValue('criticidade') || 'Baixa'));
      const dataAbertura = parseDate(getValue('dataAbertura'));
      const dataVenda = parseDate(getValue('dataVenda'));
      const notaFiscal = String(getValue('notaFiscal') || getValue('numeroNF') || '').trim();
      const pedidoCompra = String(getValue('pedidoCompra') || '').trim();
      const previsaoChegada = parseDate(getValue('previsaoChegada'));
      const ordemManutencao = String(getValue('ordemManutencao') || '').trim();
      const status = determineStatus(notaFiscal, dataVenda, pedidoCompra, ordemManutencao, quantidadeFaturada, quantidadePedida, qty);
      const diasEmAberto = dataAbertura ? Math.floor((Date.now() - dataAbertura.getTime()) / (1000 * 60 * 60 * 24)) : 0;
      const diasParaEntrega = previsaoChegada ? Math.floor((previsaoChegada.getTime() - Date.now()) / (1000 * 60 * 60 * 24)) : null;

      const record: OpportunityRecord = {
        id: globalId++,
        origemAba: sheetName,
        empresa,
        cliente,
        descricao: String(getValue('descricao') || '').trim(),
        equipamento: String(getValue('equipamento') || '').trim(),
        dataAbertura,
        mes: String(getValue('mes') || '').trim(),
        pn: String(getValue('pn') || '').trim(),
        partname: String(getValue('partname') || '').trim(),
        qty,
        criticidade,
        betim,
        lic,
        emEstoque: String(getValue('emEstoque') || '').trim(),
        importacao: parseNumber(getValue('importacao')),
        dataTroca: parseDate(getValue('dataTroca')),
        dataAberturaOM: parseDate(getValue('dataAberturaOM')),
        ordemManutencao,
        pedidoCompra,
        requisicaoCompra: String(getValue('requisicaoCompra') || '').trim(),
        previsaoChegada,
        notaFiscal,
        dataVenda,
        followUpComercial: String(getValue('followUpComercial') || '').trim(),
        followUpLocal: '',
        dataFollowUp: null,
        vinculoPasSvs: String(getValue('vinculoPasSvs') || '').trim(),
        numeroPedido: String(getValue('numeroPedido') || '').trim(),
        dataRecebimentoPedido: parseDate(getValue('dataRecebimentoPedido')),
        tipoPedido: String(getValue('tipoPedido') || '').trim(),
        dataEntregaSolicitada: parseDate(getValue('dataEntregaSolicitada')),
        partNumber: String(getValue('partNumber') || '').trim(),
        replace: String(getValue('replace') || '').trim(),
        descricao1: String(getValue('descricao1') || '').trim(),
        quantidadePedida: parseNumber(getValue('quantidadePedida')),
        disponibilidade: String(getValue('disponibilidade') || '').trim(),
        quantidadeFaturada: parseNumber(getValue('quantidadeFaturada')),
        numeroCigam: String(getValue('numeroCigam') || '').trim(),
        numeroProcessoImportacao: String(getValue('numeroProcessoImportacao') || '').trim(),
        numeroNF: String(getValue('numeroNF') || '').trim(),
        dataEmissaoNF: parseDate(getValue('dataEmissaoNF')),
        observacao: String(getValue('observacao') || '').trim(),
        status,
        diasEmAberto,
        diasParaEntrega,
        estoqueDisponivel: lic + betim,
      };

      if (empresa || cliente || record.pn) {
        allParsedData.push(record);
      }
    }
  }

  allParsedData.sort((a, b) => (b.dataAbertura?.getTime() || 0) - (a.dataAbertura?.getTime() || 0));
  return allParsedData;
}

export function formatDateBR(date: Date | null): string {
  if (!date) return '';
  return date.toLocaleDateString('pt-BR');
}
