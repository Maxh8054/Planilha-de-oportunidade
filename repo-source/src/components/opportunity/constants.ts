// Planilha de Oportunidade - Constants

import type { EstoqueFilterType } from './types';

export const SHEET_NAMES = ['PAS SVS UEM', 'PAS SVS RED', 'PAS SVS Lundin'];

export const SHEET_DISPLAY_NAMES: Record<string, string> = {
  'PAS SVS UEM': 'U&M',
  'PAS SVS RED': 'R&D',
  'PAS SVS Lundin': 'Lundin',
};

export const COLORS = {
  primary: '#6c757d',
  secondary: '#868e96',
  success: '#28a745',
  warning: '#ffc107',
  danger: '#dc3545',
  info: '#17a2b8',
  lic: '#10b981',
  betim: '#3b82f6',
};

export const CRITICITY_COLORS: Record<string, string> = {
  'Baixa': '#28a745',
  'Média': '#ffc107',
  'Alta': '#dc3545',
};

export const STATUS_CONFIG: Record<string, { color: string; label: string }> = {
  'sem_om': { color: '#dc3545', label: 'Sem OM' },
  'com_om': { color: '#fd7e14', label: 'Com OM' },
  'com_po_sem_faturamento': { color: '#ffc107', label: 'Com PO' },
  'faturado_parcial': { color: '#0ea5e9', label: 'Faturado Parcial' },
  'vendido': { color: '#28a745', label: 'Vendido' },
};

export const PIE_COLORS = ['#6c757d', '#28a745', '#ffc107', '#dc3545', '#17a2b8', '#868e96'];

export const MONTHS = [
  { value: '1', label: 'Janeiro' },
  { value: '2', label: 'Fevereiro' },
  { value: '3', label: 'Março' },
  { value: '4', label: 'Abril' },
  { value: '5', label: 'Maio' },
  { value: '6', label: 'Junho' },
  { value: '7', label: 'Julho' },
  { value: '8', label: 'Agosto' },
  { value: '9', label: 'Setembro' },
  { value: '10', label: 'Outubro' },
  { value: '11', label: 'Novembro' },
  { value: '12', label: 'Dezembro' },
];

export const COLUMN_MAP: Record<string, string> = {
  'Empresa': 'empresa',
  'Cliente': 'cliente',
  'DESCRIÇÃO': 'descricao',
  'DESCRICAO': 'descricao',
  'EQUIPAMENTO': 'equipamento',
  'DATA ABERTURA': 'dataAbertura',
  'Mês': 'mes',
  'Mes': 'mes',
  'PN': 'pn',
  'Partname': 'partname',
  'QTY': 'qty',
  'CRITICIDADE': 'criticidade',
  'Betim': 'betim',
  'LIC': 'lic',
  'EM ESTOQUE': 'emEstoque',
  'Importação': 'importacao',
  'Importacao': 'importacao',
  'Data de troca': 'dataTroca',
  'Data Abertura OM': 'dataAberturaOM',
  'ORDEM DE MANUTENÇÃO': 'ordemManutencao',
  'ORDEM DE MANUTENCAO': 'ordemManutencao',
  'PEDIDO DE COMPRA': 'pedidoCompra',
  'REQUISIÇÃO DE COMPRA': 'requisicaoCompra',
  'REQUISICAO DE COMPRA': 'requisicaoCompra',
  'PREVISÃO DE CHEGADA': 'previsaoChegada',
  'PREVISAO DE CHEGADA': 'previsaoChegada',
  'NOTA FISCAL': 'notaFiscal',
  'DAT.Venda': 'dataVenda',
  'Follow Up/comercial': 'followUpComercial',
  'VÍNCULO PAS SVS': 'vinculoPasSvs',
  'VINCULO PAS SVS': 'vinculoPasSvs',
  'Número do pedido': 'numeroPedido',
  'Numero do pedido': 'numeroPedido',
  'Data de recebimento do pedido': 'dataRecebimentoPedido',
  'Tipo do pedido': 'tipoPedido',
  'Data de entrega solicitada': 'dataEntregaSolicitada',
  'Part number': 'partNumber',
  'Part Number': 'partNumber',
  'Replace': 'replace',
  'Descrição.1': 'descricao1',
  'Descricao.1': 'descricao1',
  'Quantidade pedida': 'quantidadePedida',
  'Disponibilidade': 'disponibilidade',
  'Quantidade faturada': 'quantidadeFaturada',
  'Número do CIGAM': 'numeroCigam',
  'Numero do CIGAM': 'numeroCigam',
  'Número do processo importação': 'numeroProcessoImportacao',
  'Numero do processo importacao': 'numeroProcessoImportacao',
  'Número da NF': 'numeroNF',
  'Numero da NF': 'numeroNF',
  'Data de emissão da NF': 'dataEmissaoNF',
  'Data de emissao da NF': 'dataEmissaoNF',
  'Observação': 'observacao',
  'Observacao': 'observacao',
};

export const STATUS_OPTIONS = ['sem_om', 'com_om', 'com_po_sem_faturamento', 'faturado_parcial', 'vendido'] as const;

export const PERIOD_OPTIONS: { value: string; label: string }[] = [
  { value: 'all', label: 'Todos' },
  { value: '7d', label: 'Últimos 7 dias' },
  { value: '30d', label: 'Últimos 30 dias' },
  { value: '90d', label: 'Últimos 90 dias' },
  { value: '2024', label: '2024' },
  { value: '2025', label: '2025' },
  { value: '2026', label: '2026' },
];

export const ESTOQUE_OPTIONS: { value: EstoqueFilterType; label: string }[] = [
  { value: 'all', label: 'Todos' },
  { value: 'lic', label: 'Com estoque LIC' },
  { value: 'betim', label: 'Com estoque Betim' },
];

export const DIAS_OPTIONS: { value: string; label: string }[] = [
  { value: 'all', label: 'Todos' },
  { value: '<30', label: '< 30 dias' },
  { value: '30-60', label: '30-60 dias' },
  { value: '>60', label: '> 60 dias' },
  { value: '>90', label: '> 90 dias' },
];

export const ANALISE_OPTIONS: { value: string; label: string }[] = [
  { value: 'all', label: 'Todos' },
  { value: 'completos', label: 'Completos' },
  { value: 'incompletos', label: 'Incompletos' },
  { value: 'com_followup', label: 'Com Follow Up' },
  { value: 'sem_followup', label: 'Sem Follow Up' },
];

export const PRAZO_OPTIONS: { value: string; label: string }[] = [
  { value: 'all', label: 'Todos' },
  { value: 'atrasados', label: 'Atrasados' },
  { value: 'este_mes', label: 'Este Mês' },
  { value: 'futuro', label: 'Futuro' },
];
