'use client';

import { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  Card,
  CardContent,
  CardHeader,
  CardTitle,
} from '@/components/ui/card';
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from '@/components/ui/select';
import { Badge } from '@/components/ui/badge';
import { Button } from '@/components/ui/button';
import { Input } from '@/components/ui/input';
import { Tabs, TabsContent, TabsList, TabsTrigger } from '@/components/ui/tabs';
import {
  Table,
  TableBody,
  TableCell,
  TableHead,
  TableHeader,
  TableRow,
} from '@/components/ui/table';
import {
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
  ResponsiveContainer,
  BarChart,
  Bar,
  PieChart,
  Pie,
  Cell,
} from 'recharts';
import {
  Upload,
  FileSpreadsheet,
  Download,
  ShoppingCart,
  BarChart3,
  Target,
  Package,
  Wrench,
  TrendingUp,
  Filter,
  ChevronDown,
  X,
  MessageSquare,
  Plus,
  MessageCircle,
  FileJson,
  Clock,
  CheckCircle,
  RefreshCw,
} from 'lucide-react';
import { PasswordModal } from '@/components/PasswordModal';
import { Dialog, DialogContent, DialogDescription, DialogFooter, DialogHeader, DialogTitle } from '@/components/ui/dialog';
import { Textarea } from '@/components/ui/textarea';
import { Checkbox } from '@/components/ui/checkbox';
import { DropdownMenu, DropdownMenuContent, DropdownMenuItem, DropdownMenuTrigger } from '@/components/ui/dropdown-menu';
import { Collapsible, CollapsibleContent, CollapsibleTrigger } from '@/components/ui/collapsible';
import { FilterDropdown } from '@/components/FilterDropdown';

import type { OpportunityRecord } from '@/components/opportunity/types';
import {
  SHEET_NAMES,
  SHEET_DISPLAY_NAMES,
  COLORS,
  CRITICITY_COLORS,
  STATUS_CONFIG,
  PIE_COLORS,
  MONTHS,
  COLUMN_MAP,
} from '@/components/opportunity/constants';
import {
  parseDate,
  parseNumber,
  parseCriticidade,
  determineStatus,
  getStockInfo,
  formatDateBR,
} from '@/components/opportunity/utils';

// Stock Badge Component
function AnaliseEstoqueBadge({ record }: { record: OpportunityRecord }) {
  const info = getStockInfo(record);
  return (
    <div className="flex flex-col items-center gap-1">
      <Badge style={{ backgroundColor: info.bgColor, color: info.textColor }}>
        {info.localEstoque}
      </Badge>
      <span className="text-xs text-slate-500">
        {info.coberturaTotal ? '✓' : `Falta: ${info.saldo}`}
      </span>
    </div>
  );
}

export default function SalesOpportunityDashboard() {
  const [data, setData] = useState<OpportunityRecord[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const [isSyncing, setIsSyncing] = useState(false);
  const [lastSync, setLastSync] = useState<string | null>(null);
  const [showPasswordModal, setShowPasswordModal] = useState(false);
  const [pendingFile, setPendingFile] = useState<File | null>(null);
  const [activeTab, setActiveTab] = useState('overview');
  const fileInputRef = useRef<HTMLInputElement>(null);
  const hasLoadedRef = useRef(false);

  // Helper: parse dates from JSON
  const parseDates = (d: Record<string, unknown>, index: number): Record<string, unknown> => ({
    ...d,
    id: index + 1,
    dataAbertura: d.dataAbertura ? new Date(d.dataAbertura as string) : null,
    dataTroca: d.dataTroca ? new Date(d.dataTroca as string) : null,
    dataAberturaOM: d.dataAberturaOM ? new Date(d.dataAberturaOM as string) : null,
    previsaoChegada: d.previsaoChegada ? new Date(d.previsaoChegada as string) : null,
    dataVenda: d.dataVenda ? new Date(d.dataVenda as string) : null,
    dataFollowUp: d.dataFollowUp ? new Date(d.dataFollowUp as string) : null,
    dataRecebimentoPedido: d.dataRecebimentoPedido ? new Date(d.dataRecebimentoPedido as string) : null,
    dataEntregaSolicitada: d.dataEntregaSolicitada ? new Date(d.dataEntregaSolicitada as string) : null,
    dataEmissaoNF: d.dataEmissaoNF ? new Date(d.dataEmissaoNF as string) : null,
  });

  // Helper: serialize dates to JSON
  const serializeData = (records: OpportunityRecord[]) => records.map(d => ({
    ...d,
    dataAbertura: d.dataAbertura?.toISOString() || null,
    dataTroca: d.dataTroca?.toISOString() || null,
    dataAberturaOM: d.dataAberturaOM?.toISOString() || null,
    previsaoChegada: d.previsaoChegada?.toISOString() || null,
    dataVenda: d.dataVenda?.toISOString() || null,
    dataFollowUp: d.dataFollowUp?.toISOString() || null,
    dataRecebimentoPedido: d.dataRecebimentoPedido?.toISOString() || null,
    dataEntregaSolicitada: d.dataEntregaSolicitada?.toISOString() || null,
    dataEmissaoNF: d.dataEmissaoNF?.toISOString() || null,
  }));

  // Load from API on mount
  useEffect(() => {
    if (hasLoadedRef.current) return;
    hasLoadedRef.current = true;
    loadData();
  }, []);

  async function loadData() {
    setIsSyncing(true);
    try {
      const res = await fetch('/api/dashboard-data');
      const json = await res.json();
      if (json.data && Array.isArray(json.data) && json.data.length > 0) {
        const parsed = json.data.map((d: Record<string, unknown>, i: number) => parseDates(d, i) as OpportunityRecord);
        setData(parsed);
        setLastSync(json.updatedAt ? new Date(json.updatedAt).toLocaleString('pt-BR') : new Date().toLocaleString('pt-BR'));
      } else {
        // Fallback to localStorage
        try {
          const saved = localStorage.getItem('dashboard_oportunidades_data');
          if (saved) {
            const parsed = JSON.parse(saved);
            if (parsed.data && Array.isArray(parsed.data)) {
              setData(parsed.data.map((d: Record<string, unknown>, i: number) => parseDates(d, i) as OpportunityRecord));
            }
          }
        } catch { /* ignore */ }
      }
    } catch {
      // API not available (local dev) - fallback to localStorage
      try {
        const saved = localStorage.getItem('dashboard_oportunidades_data');
        if (saved) {
          const parsed = JSON.parse(saved);
          if (parsed.data && Array.isArray(parsed.data)) {
            setData(parsed.data.map((d: Record<string, unknown>, i: number) => parseDates(d, i) as OpportunityRecord));
          }
        }
      } catch { /* ignore */ }
    } finally {
      setIsSyncing(false);
    }
  }

  // Save to API + localStorage whenever data changes
  useEffect(() => {
    if (data.length === 0) return;
    // Always save to localStorage
    const exportData = {
      exportDate: new Date().toISOString(),
      totalRecords: data.length,
      data: serializeData(data),
    };
    localStorage.setItem('dashboard_oportunidades_data', JSON.stringify(exportData));
    // Try to save to API (don't await, fire and forget)
    fetch('/api/dashboard-data', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ data: serializeData(data), totalRecords: data.length }),
    }).then(res => res.json()).then(json => {
      if (json.success) {
        setLastSync(new Date().toLocaleString('pt-BR'));
      } else {
        console.error('Sync failed:', json.error);
      }
    }).catch(err => {
      console.error('Sync failed:', err);
    });
  }, [data]);

  // Overview Filters
  const [empresaFilter, setEmpresaFilter] = useState<string[]>([]);
  const [clienteFilter, setClienteFilter] = useState<string[]>([]);
  const [equipamentoFilter, setEquipamentoFilter] = useState<string[]>([]);
  const [criticidadeFilter, setCriticidadeFilter] = useState<string[]>([]);
  const [statusFilter, setStatusFilter] = useState<string[]>([]);
  const [periodFilter, setPeriodFilter] = useState<string[]>([]);
  const [searchTerm, setSearchTerm] = useState('');
  const [estoqueFilter, setEstoqueFilter] = useState<string[]>([]);
  const [rankingType, setRankingType] = useState<'inspecao' | 'venda'>('inspecao');
  const [chartYearFilter, setChartYearFilter] = useState<string>('');
  const [showOverviewFilters, setShowOverviewFilters] = useState(false);
  const [showOppFilters, setShowOppFilters] = useState(false);

  // Opportunities Tab Filters (multi-select)
  const [oppMonthFilter, setOppMonthFilter] = useState<string[]>([]);
  const [oppYearFilter, setOppYearFilter] = useState<string[]>([]);
  const [oppClienteFilter, setOppClienteFilter] = useState<string[]>([]);
  const [oppStatusFilter, setOppStatusFilter] = useState<string[]>([]);
  const [oppCriticidadeFilter, setOppCriticidadeFilter] = useState<string[]>([]);
  const [oppDiasFilter, setOppDiasFilter] = useState<string[]>([]);
  const [oppAnaliseFilter, setOppAnaliseFilter] = useState<string[]>([]);
  const [oppPrazoFilter, setOppPrazoFilter] = useState<string[]>([]);
  const [oppSearchTerm, setOppSearchTerm] = useState('');
  const [oppOrigemFilter, setOppOrigemFilter] = useState<string[]>([]);

  // Follow Up Modal State
  const [showFollowUpModal, setShowFollowUpModal] = useState(false);
  const [followUpRecord, setFollowUpRecord] = useState<OpportunityRecord | null>(null);
  const [followUpText, setFollowUpText] = useState('');
  const [applyToAllPedido, setApplyToAllPedido] = useState(false);

  // JSON import state
  const jsonFileInputRef = useRef<HTMLInputElement>(null);
  const [showJsonImportModal, setShowJsonImportModal] = useState(false);
  const [pendingJsonFile, setPendingJsonFile] = useState<File | null>(null);

  // Available filter options
  const empresas = useMemo(() => [...new Set(data.map(d => d.empresa).filter(Boolean))].sort(), [data]);
  const clientes = useMemo(() => [...new Set(data.map(d => d.cliente).filter(Boolean))].sort(), [data]);
  const equipamentos = useMemo(() => [...new Set(data.map(d => d.equipamento).filter(Boolean))].sort(), [data]);
  const criticidades = useMemo(() => [...new Set(data.map(d => d.criticidade).filter(Boolean))].sort(), [data]);
  const origens = useMemo(() => [...new Set(data.map(d => d.origemAba).filter(Boolean))].sort(), [data]);

  // Memoized option arrays for FilterDropdowns
  const empresaOptions = useMemo(() => empresas.map(e => ({ value: e, label: e })), [empresas]);
  const clienteOptions = useMemo(() => clientes.map(c => ({ value: c, label: c })), [clientes]);
  const equipamentoOptions = useMemo(() => equipamentos.map(e => ({ value: e, label: e })), [equipamentos]);
  const criticidadeOptions = useMemo(() => criticidades.map(c => ({ value: c, label: c })), [criticidades]);
  const statusOptions = useMemo(() => Object.entries(STATUS_CONFIG).map(([k, v]) => ({ value: k, label: v.label })), []);
  const origemOptions = useMemo(() => origens.map(o => ({ value: o, label: o.replace('PAS SVS ', '') })), [origens]);
  const monthOptions = useMemo(() => MONTHS.map(m => ({ value: m.value, label: m.label })), []);
  const diasOptions = useMemo(() => [
    { value: '<30', label: '< 30 dias' },
    { value: '30-60', label: '30-60 dias' },
    { value: '>60', label: '> 60 dias' },
    { value: '>90', label: '> 90 dias' },
  ] as const, []);
  const analiseOptions = useMemo(() => [
    { value: 'completos', label: 'Completos' },
    { value: 'incompletos', label: 'Incompletos' },
    { value: 'com_followup', label: 'Com Follow Up' },
    { value: 'sem_followup', label: 'Sem Follow Up' },
  ] as const, []);
  const prazoOptions = useMemo(() => [
    { value: 'atrasados', label: 'Atrasados' },
    { value: 'este_mes', label: 'Este Mês' },
    { value: 'futuro', label: 'Futuro' },
  ] as const, []);
  const periodoOptions = useMemo(() => [
    { value: '7d', label: 'Últimos 7 dias' },
    { value: '30d', label: 'Últimos 30 dias' },
    { value: '90d', label: 'Últimos 90 dias' },
    { value: '2024', label: '2024' },
    { value: '2025', label: '2025' },
    { value: '2026', label: '2026' },
  ] as const, []);
  const estoqueOptions = useMemo(() => [
    { value: 'lic', label: 'Com LIC' },
    { value: 'betim', label: 'Com Betim' },
  ] as const, []);

  const availableYears = useMemo(() => {
    const years = new Set<number>();
    data.forEach(d => {
      if (d.dataAbertura) years.add(d.dataAbertura.getFullYear());
    });
    return Array.from(years).sort().map(y => y.toString());
  }, [data]);

  const yearOptions = useMemo(() => availableYears.map(y => ({ value: y, label: y })), [availableYears]);

  // Auto-select most recent year for chart
  const effectiveChartYear = chartYearFilter || (availableYears.length > 0 ? availableYears[availableYears.length - 1] : '');
  useEffect(() => {
    if (!chartYearFilter && availableYears.length > 0) {
      setChartYearFilter(availableYears[availableYears.length - 1]);
    }
  }, [availableYears]);

  // Overview filtered data
  const filteredData = useMemo(() => {
    let result = data;
    if (searchTerm) {
      const search = searchTerm.toLowerCase();
      result = result.filter(d =>
        d.pn?.toLowerCase().includes(search) ||
        d.partname?.toLowerCase().includes(search) ||
        d.cliente?.toLowerCase().includes(search) ||
        d.descricao?.toLowerCase().includes(search) ||
        d.ordemManutencao?.toLowerCase().includes(search) ||
        d.pedidoCompra?.toLowerCase().includes(search)
      );
    }
    if (empresaFilter.length > 0) result = result.filter(d => empresaFilter.includes(d.empresa));
    if (clienteFilter.length > 0) result = result.filter(d => clienteFilter.includes(d.cliente));
    if (equipamentoFilter.length > 0) result = result.filter(d => equipamentoFilter.includes(d.equipamento));
    if (criticidadeFilter.length > 0) result = result.filter(d => criticidadeFilter.includes(d.criticidade));
    if (statusFilter.length > 0) result = result.filter(d => statusFilter.includes(d.status));
    if (periodFilter.length > 0) {
      const now = new Date();
      result = result.filter(d => {
        if (!d.dataAbertura) return false;
        for (const p of periodFilter) {
          const filterDate = new Date();
          switch (p) {
            case '7d': filterDate.setDate(now.getDate() - 7); if (d.dataAbertura >= filterDate) return true; break;
            case '30d': filterDate.setDate(now.getDate() - 30); if (d.dataAbertura >= filterDate) return true; break;
            case '90d': filterDate.setDate(now.getDate() - 90); if (d.dataAbertura >= filterDate) return true; break;
            case '2024': if (d.dataAbertura.getFullYear() === 2024) return true; break;
            case '2025': if (d.dataAbertura.getFullYear() === 2025) return true; break;
            case '2026': if (d.dataAbertura.getFullYear() === 2026) return true; break;
          }
        }
        return false;
      });
    }
    if (estoqueFilter.length > 0) {
      result = result.filter(d => {
        for (const e of estoqueFilter) {
          if (e === 'lic' && d.lic > 0) return true;
          if (e === 'betim' && d.betim > 0) return true;
        }
        return false;
      });
    }
    return result;
  }, [data, searchTerm, empresaFilter, clienteFilter, equipamentoFilter, criticidadeFilter, statusFilter, periodFilter, estoqueFilter]);

  // Opportunities Tab filtered data
  const oppFilteredData = useMemo(() => {
    let result = data;
    if (oppSearchTerm) {
      const search = oppSearchTerm.toLowerCase();
      result = result.filter(d =>
        d.pn?.toLowerCase().includes(search) ||
        d.partname?.toLowerCase().includes(search) ||
        d.cliente?.toLowerCase().includes(search) ||
        d.descricao?.toLowerCase().includes(search)
      );
    }
    if (oppOrigemFilter.length > 0) result = result.filter(d => oppOrigemFilter.includes(d.origemAba));
    if (oppMonthFilter.length > 0) result = result.filter(d => d.dataAbertura && oppMonthFilter.includes(String(d.dataAbertura.getMonth() + 1)));
    if (oppYearFilter.length > 0) result = result.filter(d => d.dataAbertura && oppYearFilter.includes(String(d.dataAbertura.getFullYear())));
    if (oppClienteFilter.length > 0) result = result.filter(d => oppClienteFilter.includes(d.cliente));
    if (oppStatusFilter.length > 0) result = result.filter(d => oppStatusFilter.includes(d.status));
    if (oppCriticidadeFilter.length > 0) result = result.filter(d => oppCriticidadeFilter.includes(d.criticidade));
    if (oppDiasFilter.length > 0) {
      result = result.filter(d => {
        for (const f of oppDiasFilter) {
          if (f === '<30' && d.diasEmAberto < 30) return true;
          if (f === '30-60' && d.diasEmAberto >= 30 && d.diasEmAberto <= 60) return true;
          if (f === '>60' && d.diasEmAberto > 60) return true;
          if (f === '>90' && d.diasEmAberto > 90) return true;
        }
        return false;
      });
    }
    if (oppAnaliseFilter.length > 0) {
      result = result.filter(d => {
        for (const f of oppAnaliseFilter) {
          if (f === 'completos' && d.quantidadeFaturada > 0 && d.quantidadeFaturada >= d.qty) return true;
          if (f === 'incompletos' && (d.quantidadeFaturada === 0 || d.quantidadeFaturada < d.qty)) return true;
          if (f === 'com_followup' && (d.followUpComercial || d.followUpLocal)) return true;
          if (f === 'sem_followup' && !d.followUpComercial && !d.followUpLocal) return true;
        }
        return false;
      });
    }
    if (oppPrazoFilter.length > 0) {
      const now = new Date();
      const cm = now.getMonth();
      const cy = now.getFullYear();
      result = result.filter(d => {
        for (const f of oppPrazoFilter) {
          if (f === 'atrasados' && d.dataEntregaSolicitada && d.dataEntregaSolicitada < now && d.status !== 'vendido') return true;
          if (f === 'este_mes' && d.dataEntregaSolicitada && d.dataEntregaSolicitada.getMonth() === cm && d.dataEntregaSolicitada.getFullYear() === cy) return true;
          if (f === 'futuro' && d.dataEntregaSolicitada && d.dataEntregaSolicitada > now) return true;
        }
        return false;
      });
    }
    return result;
  }, [data, oppSearchTerm, oppMonthFilter, oppYearFilter, oppClienteFilter, oppStatusFilter, oppCriticidadeFilter, oppDiasFilter, oppAnaliseFilter, oppPrazoFilter, oppOrigemFilter]);


  // KPIs
  const kpis = useMemo(() => {
    const totalLinhas = filteredData.length;
    const totalQty = filteredData.reduce((sum, d) => sum + d.qty, 0);
    const totalPedidos = new Set(filteredData.filter(d => d.pedidoCompra).map(d => d.pedidoCompra)).size;
    const totalOMs = new Set(filteredData.filter(d => d.ordemManutencao).map(d => d.ordemManutencao)).size;
    return { totalLinhas, totalQty, totalPedidos, totalOMs };
  }, [filteredData]);

  // Opportunities KPIs
  const oppKpis = useMemo(() => {
    const totalLinhas = oppFilteredData.length;
    const totalQty = oppFilteredData.reduce((sum, d) => sum + d.qty, 0);
    const totalPedidos = new Set(oppFilteredData.filter(d => d.pedidoCompra).map(d => d.pedidoCompra)).size;
    const totalOMs = new Set(oppFilteredData.filter(d => d.ordemManutencao).map(d => d.ordemManutencao)).size;
    const linhasFiltradas = oppFilteredData.length;
    const totalEmAberto = oppFilteredData.filter(d => d.status !== 'vendido' && d.status !== 'faturado_parcial').length;
    const totalVendido = oppFilteredData.reduce((sum, d) => sum + d.quantidadeFaturada, 0);
    const totalNF = new Set(oppFilteredData.filter(d => d.notaFiscal).map(d => d.notaFiscal)).size;
    const faturadoParcial = oppFilteredData.filter(d => d.status === 'faturado_parcial').length;
    return { totalLinhas, totalQty, totalPedidos, totalOMs, linhasFiltradas, totalEmAberto, totalVendido, totalNF, faturadoParcial };
  }, [oppFilteredData]);

  // Charts data
  const clientChartData = useMemo(() => {
    const clientData: Record<string, { qty: number; vendido: number }> = {};
    filteredData.forEach(d => {
      if (d.cliente) {
        if (!clientData[d.cliente]) clientData[d.cliente] = { qty: 0, vendido: 0 };
        clientData[d.cliente].qty += d.qty;
        if (d.status === 'vendido') clientData[d.cliente].vendido += d.qty;
      }
    });
    return Object.entries(clientData)
      .sort((a, b) => b[1].qty - a[1].qty)
      .slice(0, 10)
      .map(([name, d]) => ({ name, Quantidade: d.qty, Vendido: d.vendido }));
  }, [filteredData]);

  const equipmentChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => { if (d.equipamento) counts[d.equipamento] = (counts[d.equipamento] || 0) + d.qty; });
    return Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 10).map(([name, value]) => ({ name, value }));
  }, [filteredData]);

  const criticityChartData = useMemo(() => {
    const counts: Record<string, number> = { 'Baixa': 0, 'Média': 0, 'Alta': 0 };
    filteredData.forEach(d => {
      if (d.criticidade === 'Alta') counts['Alta']++;
      else if (d.criticidade === 'Média') counts['Média']++;
      else counts['Baixa']++;
    });
    return [
      { name: 'Baixa', value: counts['Baixa'], color: CRITICITY_COLORS['Baixa'] },
      { name: 'Média', value: counts['Média'], color: CRITICITY_COLORS['Média'] },
      { name: 'Alta', value: counts['Alta'], color: CRITICITY_COLORS['Alta'] },
    ].filter(d => d.value > 0);
  }, [filteredData]);

  const statusChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    filteredData.forEach(d => { counts[d.status] = (counts[d.status] || 0) + 1; });
    return Object.entries(counts).map(([name, value]) => ({ name: STATUS_CONFIG[name]?.label || name, value }));
  }, [filteredData]);

  const monthlyChartData = useMemo(() => {
    const monthCounts: Record<string, number> = {};
    const monthNames = ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'];
    filteredData.forEach(d => {
      if (d.dataAbertura) {
        if (effectiveChartYear && d.dataAbertura.getFullYear() !== parseInt(effectiveChartYear)) return;
        const key = `${monthNames[d.dataAbertura.getMonth()]}/${d.dataAbertura.getFullYear()}`;
        monthCounts[key] = (monthCounts[key] || 0) + d.qty;
      }
    });
    return Object.entries(monthCounts)
      .sort((a, b) => {
        const [aM, aY] = a[0].split('/');
        const [bM, bY] = b[0].split('/');
        return (parseInt(aY) * 12 + monthNames.indexOf(aM)) - (parseInt(bY) * 12 + monthNames.indexOf(bM));
      })
      .map(([name, value]) => ({ name, value }));
  }, [filteredData, effectiveChartYear]);

  const origemChartData = useMemo(() => {
    const counts: Record<string, number> = {};
    data.forEach(d => {
      if (d.origemAba) { const label = d.origemAba.replace('PAS SVS ', ''); counts[label] = (counts[label] || 0) + 1; }
    });
    return Object.entries(counts).map(([name, value]) => ({ name, value }));
  }, [data]);

  // Ranking data - only by equipamento (column D)
  const rankingData = useMemo(() => {
    const items: Record<string, { inspecao: number; venda: number }> = {};
    filteredData.forEach(d => {
      if (!d.equipamento) return;
      const name = d.equipamento;
      if (!items[name]) items[name] = { inspecao: 0, venda: 0 };
      items[name].inspecao += d.qty;
      items[name].venda += d.quantidadeFaturada;
    });
    return Object.entries(items)
      .sort((a, b) => b[1][rankingType] - a[1][rankingType])
      .map(([name, counts], i) => ({ position: i + 1, name, inspecao: counts.inspecao, venda: counts.venda }));
  }, [filteredData, rankingType]);

  // File handling
  const handleFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setPendingFile(file);
    setShowPasswordModal(true);
    if (fileInputRef.current) fileInputRef.current.value = '';
  };

  const handleImportConfirmed = async () => {
    if (!pendingFile) return;
    setIsLoading(true);
    try {
      const fileData = await pendingFile.arrayBuffer();
      const workbook = XLSX.read(fileData);
      const allParsedData: OpportunityRecord[] = [];
      let globalId = 1;

      for (const sheetName of SHEET_NAMES) {
        const actualSheetName = workbook.SheetNames.find(n => n.toLowerCase() === sheetName.toLowerCase());
        if (!actualSheetName) { console.warn(`Aba "${sheetName}" não encontrada`); continue; }
        const worksheet = workbook.Sheets[actualSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as unknown[][];
        if (jsonData.length < 2) continue;
        const headers = jsonData[0] as string[];
        const columnIndices: Record<string, number> = {};
        headers.forEach((header, index) => {
          const mappedKey = COLUMN_MAP[header?.toString().trim()];
          if (mappedKey) columnIndices[mappedKey] = index;
        });

        for (let i = 1; i < jsonData.length; i++) {
          const row = jsonData[i] as unknown[];
          if (!row || row.length === 0) continue;
          const getValue = (key: string): unknown => { const idx = columnIndices[key]; return idx !== undefined ? row[idx] : undefined; };
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
          const quantidadePedida = parseNumber(getValue('quantidadePedida'));
          const quantidadeFaturada = parseNumber(getValue('quantidadeFaturada'));
          const status = determineStatus(notaFiscal, dataVenda, pedidoCompra, ordemManutencao, quantidadeFaturada, quantidadePedida, qty);
          const diasEmAberto = dataAbertura ? Math.floor((Date.now() - dataAbertura.getTime()) / (1000 * 60 * 60 * 24)) : 0;
          const diasParaEntrega = previsaoChegada ? Math.floor((previsaoChegada.getTime() - Date.now()) / (1000 * 60 * 60 * 24)) : null;

          const record: OpportunityRecord = {
            id: globalId++, origemAba: sheetName, empresa, cliente,
            descricao: String(getValue('descricao') || '').trim(),
            equipamento: String(getValue('equipamento') || '').trim(),
            dataAbertura, mes: String(getValue('mes') || '').trim(),
            pn: String(getValue('pn') || '').trim(),
            partname: String(getValue('partname') || '').trim(),
            qty, criticidade, betim, lic,
            emEstoque: String(getValue('emEstoque') || '').trim(),
            importacao: parseNumber(getValue('importacao')),
            dataTroca: parseDate(getValue('dataTroca')),
            dataAberturaOM: parseDate(getValue('dataAberturaOM')),
            ordemManutencao, pedidoCompra,
            requisicaoCompra: String(getValue('requisicaoCompra') || '').trim(),
            previsaoChegada, notaFiscal, dataVenda,
            followUpComercial: String(getValue('followUpComercial') || '').trim(),
            followUpLocal: '', dataFollowUp: null,
            vinculoPasSvs: String(getValue('vinculoPasSvs') || '').trim(),
            numeroPedido: String(getValue('numeroPedido') || '').trim(),
            dataRecebimentoPedido: parseDate(getValue('dataRecebimentoPedido')),
            tipoPedido: String(getValue('tipoPedido') || '').trim(),
            dataEntregaSolicitada: parseDate(getValue('dataEntregaSolicitada')),
            partNumber: String(getValue('partNumber') || '').trim(),
            replace: String(getValue('replace') || '').trim(),
            descricao1: String(getValue('descricao1') || '').trim(),
            quantidadePedida,
            disponibilidade: String(getValue('disponibilidade') || '').trim(),
            quantidadeFaturada,
            numeroCigam: String(getValue('numeroCigam') || '').trim(),
            numeroProcessoImportacao: String(getValue('numeroProcessoImportacao') || '').trim(),
            numeroNF: String(getValue('numeroNF') || '').trim(),
            dataEmissaoNF: parseDate(getValue('dataEmissaoNF')),
            observacao: String(getValue('observacao') || '').trim(),
            status, diasEmAberto, diasParaEntrega,
            estoqueDisponivel: lic + betim,
          };
          if (empresa || cliente || record.pn) allParsedData.push(record);
        }
      }
      allParsedData.sort((a, b) => (b.dataAbertura?.getTime() || 0) - (a.dataAbertura?.getTime() || 0));
      setData(allParsedData);
      setPendingFile(null);
      alert(`Importação concluída! ${allParsedData.length} registros carregados de ${SHEET_NAMES.filter(s => workbook.SheetNames.some(n => n.toLowerCase() === s.toLowerCase())).length} abas.`);
    } catch (error) {
      console.error('Erro ao importar:', error);
      alert('Erro ao importar arquivo! Verifique o formato.');
    } finally { setIsLoading(false); }
  };

  const exportToExcel = (tipo: 'completo' | 'servicos') => {
    if (data.length === 0) return;
    const wb = XLSX.utils.book_new();
    const dataByOrigin: Record<string, OpportunityRecord[]> = {};
    data.forEach(d => { if (!dataByOrigin[d.origemAba]) dataByOrigin[d.origemAba] = []; dataByOrigin[d.origemAba].push(d); });

    SHEET_NAMES.forEach(sheetName => {
      const sheetData = dataByOrigin[sheetName] || [];
      let exportData: Record<string, unknown>[];
      if (tipo === 'servicos') {
        exportData = sheetData.map(d => ({
          'Empresa': d.empresa, 'Cliente': d.cliente, 'DESCRIÇÃO': d.descricao,
          'EQUIPAMENTO': d.equipamento, 'DATA ABERTURA': formatDateBR(d.dataAbertura),
          'Mês': d.mes, 'PN': d.pn, 'Partname': d.partname, 'QTY': d.qty,
          'CRITICIDADE': d.criticidade, 'EM ESTOQUE': d.emEstoque,
          'Data de troca': formatDateBR(d.dataTroca),
          'Data Abertura OM': formatDateBR(d.dataAberturaOM),
          'ORDEM DE MANUTENÇÃO': d.ordemManutencao, 'PEDIDO DE COMPRA': d.pedidoCompra,
          'REQUISIÇÃO DE COMPRA': d.requisicaoCompra,
          'PREVISÃO DE CHEGADA': formatDateBR(d.previsaoChegada),
          'NOTA FISCAL': d.notaFiscal, 'DAT.Venda': formatDateBR(d.dataVenda),
          'Follow Up/comercial': d.followUpComercial, 'Follow Up Local': d.followUpLocal,
        }));
      } else {
        exportData = sheetData.map(d => ({
          'Empresa': d.empresa, 'Cliente': d.cliente, 'DESCRIÇÃO': d.descricao,
          'EQUIPAMENTO': d.equipamento, 'DATA ABERTURA': formatDateBR(d.dataAbertura),
          'Mês': d.mes, 'PN': d.pn, 'Partname': d.partname, 'QTY': d.qty,
          'Quantidade Inspeção': d.qty, 'Quantidade pedida': d.quantidadePedida,
          'CRITICIDADE': d.criticidade, 'Betim': d.betim, 'LIC': d.lic,
          'EM ESTOQUE': d.emEstoque, 'Importação': d.importacao,
          'Data de troca': formatDateBR(d.dataTroca),
          'Data Abertura OM': formatDateBR(d.dataAberturaOM),
          'ORDEM DE MANUTENÇÃO': d.ordemManutencao, 'PEDIDO DE COMPRA': d.pedidoCompra,
          'REQUISIÇÃO DE COMPRA': d.requisicaoCompra,
          'PREVISÃO DE CHEGADA': formatDateBR(d.previsaoChegada),
          'NOTA FISCAL': d.notaFiscal, 'DAT.Venda': formatDateBR(d.dataVenda),
          'Follow Up/comercial': d.followUpComercial, 'Follow Up Local': d.followUpLocal,
          'Data Follow Up': formatDateBR(d.dataFollowUp),
          'VÍNCULO PAS SVS': d.vinculoPasSvs, 'Número do pedido': d.numeroPedido,
          'Data de recebimento do pedido': formatDateBR(d.dataRecebimentoPedido),
          'Tipo do pedido': d.tipoPedido,
          'Data de entrega solicitada': formatDateBR(d.dataEntregaSolicitada),
          'Part number': d.partNumber, 'Replace': d.replace, 'Descrição.1': d.descricao1,
          'Disponibilidade': d.disponibilidade, 'Quantidade faturada': d.quantidadeFaturada,
          'Número do CIGAM': d.numeroCigam,
          'Número do processo importação': d.numeroProcessoImportacao,
          'Número da NF': d.numeroNF, 'Data de emissão da NF': formatDateBR(d.dataEmissaoNF),
          'Observação': d.observacao,
        }));
      }
      const ws = XLSX.utils.json_to_sheet(exportData);
      const exportSheetName = tipo === 'servicos' ? (SHEET_DISPLAY_NAMES[sheetName] || sheetName) : sheetName;
      XLSX.utils.book_append_sheet(wb, ws, exportSheetName);
    });
    const suffix = tipo === 'servicos' ? '_servicos' : '_completo';
    XLSX.writeFile(wb, `oportunidades_export${suffix}_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const exportToJSON = () => {
    if (data.length === 0) { alert('Nenhum dado para exportar!'); return; }
    const exportData = {
      exportDate: new Date().toISOString(),
      totalRecords: data.length,
      data: data.map(d => ({
        ...d,
        dataAbertura: d.dataAbertura?.toISOString() || null,
        dataTroca: d.dataTroca?.toISOString() || null,
        dataAberturaOM: d.dataAberturaOM?.toISOString() || null,
        previsaoChegada: d.previsaoChegada?.toISOString() || null,
        dataVenda: d.dataVenda?.toISOString() || null,
        dataFollowUp: d.dataFollowUp?.toISOString() || null,
        dataRecebimentoPedido: d.dataRecebimentoPedido?.toISOString() || null,
        dataEntregaSolicitada: d.dataEntregaSolicitada?.toISOString() || null,
        dataEmissaoNF: d.dataEmissaoNF?.toISOString() || null,
      }))
    };
    const blob = new Blob([JSON.stringify(exportData, null, 2)], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `dashboard_oportunidades_${new Date().toISOString().split('T')[0]}.json`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleJsonFileSelect = (event: React.ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0];
    if (!file) return;
    setPendingJsonFile(file);
    setShowJsonImportModal(true);
    if (jsonFileInputRef.current) jsonFileInputRef.current.value = '';
  };

  const handleJsonImportConfirmed = async () => {
    if (!pendingJsonFile) return;
    setIsLoading(true);
    try {
      const text = await pendingJsonFile.text();
      const jsonData = JSON.parse(text);
      if (!jsonData.data || !Array.isArray(jsonData.data)) throw new Error('Formato inválido');
      const parsedData: OpportunityRecord[] = jsonData.data.map((d: Record<string, unknown>, index: number) => ({
        id: index + 1,
        origemAba: String(d.origemAba || ''), empresa: String(d.empresa || ''),
        cliente: String(d.cliente || ''), descricao: String(d.descricao || ''),
        equipamento: String(d.equipamento || ''),
        dataAbertura: d.dataAbertura ? new Date(d.dataAbertura as string) : null,
        mes: String(d.mes || ''), pn: String(d.pn || ''),
        partname: String(d.partname || ''), qty: Number(d.qty) || 0,
        criticidade: String(d.criticidade || 'Baixa'),
        betim: Number(d.betim) || 0, lic: Number(d.lic) || 0,
        emEstoque: String(d.emEstoque || ''), importacao: Number(d.importacao) || 0,
        dataTroca: d.dataTroca ? new Date(d.dataTroca as string) : null,
        dataAberturaOM: d.dataAberturaOM ? new Date(d.dataAberturaOM as string) : null,
        ordemManutencao: String(d.ordemManutencao || ''),
        pedidoCompra: String(d.pedidoCompra || ''),
        requisicaoCompra: String(d.requisicaoCompra || ''),
        previsaoChegada: d.previsaoChegada ? new Date(d.previsaoChegada as string) : null,
        notaFiscal: String(d.notaFiscal || ''),
        dataVenda: d.dataVenda ? new Date(d.dataVenda as string) : null,
        followUpComercial: String(d.followUpComercial || ''),
        followUpLocal: String(d.followUpLocal || ''),
        dataFollowUp: d.dataFollowUp ? new Date(d.dataFollowUp as string) : null,
        vinculoPasSvs: String(d.vinculoPasSvs || ''),
        numeroPedido: String(d.numeroPedido || ''),
        dataRecebimentoPedido: d.dataRecebimentoPedido ? new Date(d.dataRecebimentoPedido as string) : null,
        tipoPedido: String(d.tipoPedido || ''),
        dataEntregaSolicitada: d.dataEntregaSolicitada ? new Date(d.dataEntregaSolicitada as string) : null,
        partNumber: String(d.partNumber || ''), replace: String(d.replace || ''),
        descricao1: String(d.descricao1 || ''),
        quantidadePedida: Number(d.quantidadePedida) || 0,
        disponibilidade: String(d.disponibilidade || ''),
        quantidadeFaturada: Number(d.quantidadeFaturada) || 0,
        numeroCigam: String(d.numeroCigam || ''),
        numeroProcessoImportacao: String(d.numeroProcessoImportacao || ''),
        numeroNF: String(d.numeroNF || ''),
        dataEmissaoNF: d.dataEmissaoNF ? new Date(d.dataEmissaoNF as string) : null,
        observacao: String(d.observacao || ''),
        status: (d.status as OpportunityRecord['status']) || 'sem_om',
        diasEmAberto: Number(d.diasEmAberto) || 0,
        diasParaEntrega: d.diasParaEntrega ? Number(d.diasParaEntrega) : null,
        estoqueDisponivel: Number(d.estoqueDisponivel) || 0,
      }));
      setData(parsedData);
      setShowJsonImportModal(false);
      setPendingJsonFile(null);
      alert(`Importação JSON concluída! ${parsedData.length} registros carregados.`);
    } catch (error) {
      console.error('Erro ao importar JSON:', error);
      alert('Erro ao importar arquivo JSON! Verifique o formato.');
    } finally { setIsLoading(false); }
  };

  const clearFilters = () => {
    setEmpresaFilter([]); setClienteFilter([]); setEquipamentoFilter([]);
    setCriticidadeFilter([]); setStatusFilter([]); setPeriodFilter([]);
    setSearchTerm(''); setEstoqueFilter([]);
  };
  const clearOppFilters = () => {
    setOppMonthFilter([]); setOppYearFilter([]); setOppClienteFilter([]);
    setOppStatusFilter([]); setOppCriticidadeFilter([]); setOppDiasFilter([]);
    setOppAnaliseFilter([]); setOppPrazoFilter([]); setOppSearchTerm('');
    setOppOrigemFilter([]);
  };
  const openFollowUpModal = (record: OpportunityRecord) => {
    setFollowUpRecord(record);
    setFollowUpText(record.followUpLocal || record.followUpComercial || '');
    setApplyToAllPedido(false);
    setShowFollowUpModal(true);
  };
  const saveFollowUp = () => {
    if (!followUpRecord || !followUpText.trim()) return;
    const now = new Date();
    setData(prevData => {
      if (applyToAllPedido && followUpRecord.pedidoCompra) {
        return prevData.map(d => d.pedidoCompra === followUpRecord.pedidoCompra
          ? { ...d, followUpLocal: followUpText, dataFollowUp: now } : d);
      }
      return prevData.map(d => d.id === followUpRecord.id
        ? { ...d, followUpLocal: followUpText, dataFollowUp: now } : d);
    });
    setShowFollowUpModal(false); setFollowUpRecord(null); setFollowUpText(''); setApplyToAllPedido(false);
  };

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 to-slate-100 flex flex-col">
      {/* Header */}
      <header className="bg-white border-b-4 border-slate-400 shadow-sm sticky top-0 z-50">
        <div className="w-full px-6 py-3 flex items-center justify-between flex-wrap gap-3">
          <div className="flex items-center gap-4">
            <div className="w-12 h-12 rounded-xl bg-slate-200 flex items-center justify-center shrink-0">
              <Target className="h-6 w-6 text-slate-600" />
            </div>
            <div>
              <h1 className="text-xl font-bold text-slate-700">Dashboard de Oportunidades</h1>
              <div className="flex items-center gap-2 text-xs text-slate-400">
                {isSyncing && (
                  <><div className="h-3 w-3 animate-spin rounded-full border-2 border-slate-300 border-t-slate-600" /> Sincronizando...</>
                )}
                {!isSyncing && lastSync && (
                  <><CheckCircle className="h-3 w-3 text-emerald-500" /> Sync: {lastSync}</>
                )}
                {!isSyncing && !lastSync && data.length > 0 && (
                  <span>Local (sem banco)</span>
                )}
              </div>
            </div>
            <Button
              variant="ghost"
              size="sm"
              className="h-8 w-8 p-0"
              onClick={() => loadData()}
              disabled={isSyncing}
              title="Sincronizar"
            >
              <RefreshCw className={`h-4 w-4 text-slate-500 ${isSyncing ? 'animate-spin' : ''}`} />
            </Button>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            <input ref={fileInputRef} type="file" accept=".xlsx,.xls" onChange={handleFileSelect} className="hidden" id="excel-file-upload" />
            <label htmlFor="excel-file-upload">
              <Button asChild disabled={isLoading} className="gap-2 bg-slate-600 hover:bg-slate-700 cursor-pointer">
                <span>
                  {isLoading ? (
                    <><div className="h-4 w-4 animate-spin rounded-full border-2 border-white border-t-transparent" /> Importando...</>
                  ) : (
                    <><Upload className="h-4 w-4" /> Importar Excel</>
                  )}
                </span>
              </Button>
            </label>

            {data.length > 0 && (
              <DropdownMenu>
                <DropdownMenuTrigger asChild>
                  <Button variant="outline" className="gap-2">
                    <Download className="h-4 w-4" /> Exportar <ChevronDown className="h-4 w-4" />
                  </Button>
                </DropdownMenuTrigger>
                <DropdownMenuContent align="end">
                  <DropdownMenuItem onClick={() => exportToExcel('completo')}>
                    <FileSpreadsheet className="h-4 w-4 mr-2" /> Exportar Completo
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={() => exportToExcel('servicos')}>
                    <Wrench className="h-4 w-4 mr-2" /> Exportar Serviços
                  </DropdownMenuItem>
                  <DropdownMenuItem onClick={exportToJSON}>
                    <FileJson className="h-4 w-4 mr-2" /> Exportar JSON
                  </DropdownMenuItem>
                </DropdownMenuContent>
              </DropdownMenu>
            )}

            <input ref={jsonFileInputRef} type="file" accept=".json" onChange={handleJsonFileSelect} className="hidden" id="json-file-upload" />
            <label htmlFor="json-file-upload">
              <Button asChild variant="outline" className="gap-2 cursor-pointer">
                <span><FileJson className="h-4 w-4" /> Importar JSON</span>
              </Button>
            </label>
          </div>
        </div>
      </header>

      {/* Modals */}
      <PasswordModal
        open={showPasswordModal}
        onOpenChange={(open) => { setShowPasswordModal(open); if (!open) setPendingFile(null); }}
        onConfirm={handleImportConfirmed}
        title="Confirmar Importação"
        description="Digite a senha para importar os dados."
      />
      <PasswordModal
        open={showJsonImportModal}
        onOpenChange={(open) => { setShowJsonImportModal(open); if (!open) setPendingJsonFile(null); }}
        onConfirm={handleJsonImportConfirmed}
        title="Importar JSON"
        description="Digite a senha para importar os dados do JSON."
      />

      {/* Follow Up Modal */}
      <Dialog open={showFollowUpModal} onOpenChange={setShowFollowUpModal}>
        <DialogContent className="max-w-lg bg-white">
          <DialogHeader>
            <DialogTitle className="flex items-center gap-2">
              <MessageSquare className="h-5 w-5 text-slate-600" /> Adicionar Follow Up
            </DialogTitle>
            <DialogDescription>
              {followUpRecord && (
                <div className="mt-2 text-sm text-slate-600">
                  <p><strong>PN:</strong> {followUpRecord.pn}</p>
                  <p><strong>Descrição:</strong> {followUpRecord.partname || followUpRecord.descricao}</p>
                  <p><strong>Pedido:</strong> {followUpRecord.pedidoCompra || 'N/A'}</p>
                  <p><strong>Origem:</strong> {followUpRecord.origemAba}</p>
                </div>
              )}
            </DialogDescription>
          </DialogHeader>
          <div className="space-y-4 py-4">
            {followUpRecord?.followUpComercial && (
              <div className="p-3 bg-amber-50 border border-amber-200 rounded-lg">
                <p className="text-xs text-amber-600 font-medium mb-1">Follow Up Comercial (Excel):</p>
                <p className="text-sm text-slate-700">{followUpRecord.followUpComercial}</p>
              </div>
            )}
            {followUpRecord?.followUpLocal && (
              <div className="p-3 bg-emerald-50 border border-emerald-200 rounded-lg">
                <p className="text-xs text-emerald-600 font-medium mb-1">Follow Up Local:</p>
                <p className="text-sm text-slate-700">{followUpRecord.followUpLocal}</p>
              </div>
            )}
            <Textarea placeholder="Digite o follow up..." value={followUpText} onChange={(e) => setFollowUpText(e.target.value)} rows={4} className="resize-none" />
            {followUpRecord?.pedidoCompra && (
              <div className="flex items-center space-x-2">
                <Checkbox id="applyAll" checked={applyToAllPedido} onCheckedChange={(checked) => setApplyToAllPedido(checked as boolean)} />
                <label htmlFor="applyAll" className="text-sm text-slate-600 cursor-pointer">
                  Aplicar a todas as linhas com o mesmo pedido ({followUpRecord.pedidoCompra})
                </label>
              </div>
            )}
          </div>
          <DialogFooter>
            <Button variant="outline" onClick={() => setShowFollowUpModal(false)}>Cancelar</Button>
            <Button onClick={saveFollowUp} disabled={!followUpText.trim()}>Salvar Follow Up</Button>
          </DialogFooter>
        </DialogContent>
      </Dialog>

      <main className="w-full px-6 py-4 flex-1">
        {data.length === 0 ? (
          <Card className="border-0 shadow-lg">
            <CardContent className="py-20 text-center max-w-xl mx-auto">
              <FileSpreadsheet className="h-20 w-20 mx-auto text-slate-300 mb-6" />
              <h3 className="text-xl font-semibold text-slate-600 mb-3">Bem-vindo ao Dashboard de Oportunidades</h3>
              <p className="text-slate-500 mb-8 leading-relaxed">
                Para começar, baixe o arquivo de oportunidades no link abaixo e importe-o na plataforma.
                O arquivo contém as planilhas de acompanhamento de todas as operações.
              </p>
              <div className="bg-slate-50 rounded-xl p-6 border border-slate-200">
                <p className="text-sm font-medium text-slate-700 mb-3">
                  Arquivo: <span className="text-slate-900 font-semibold">Acompanhamento de Oportunidades PAS SVS.xlsx</span>
                </p>
                <a
                  href="https://zaminebrasil.sharepoint.com/sites/SERVIOS-LUNDIN/Shared%20Documents/Forms/AllItems.aspx?id=%2Fsites%2FSERVIOS%2DLUNDIN%2FShared%20Documents%2FSERVI%C3%87OS%2FLundin%2FOportunidades&viewid=c13d8cad%2Da802%2D45ff%2D9ca9%2De5760b0f6790"
                  target="_blank"
                  rel="noopener noreferrer"
                  className="inline-flex items-center gap-2 bg-slate-700 hover:bg-slate-800 text-white font-medium px-6 py-3 rounded-lg transition-colors"
                >
                  <Download className="h-4 w-4" />
                  Clique aqui para baixar o arquivo
                </a>
              </div>
              <div className="mt-6 flex items-center justify-center gap-3">
                <div className="h-px bg-slate-200 flex-1" />
                <span className="text-xs text-slate-400 uppercase">ou</span>
                <div className="h-px bg-slate-200 flex-1" />
              </div>
              <p className="text-sm text-slate-400 mt-4">Depois de baixar, clique em <strong>Importar Excel</strong> no topo da página</p>
            </CardContent>
          </Card>
        ) : (
          <Tabs value={activeTab} onValueChange={setActiveTab} className="space-y-4">
            <TabsList className="bg-white border shadow-sm">
              <TabsTrigger value="overview" className="gap-2">
                <BarChart3 className="h-4 w-4" /> Visão Geral
              </TabsTrigger>
              <TabsTrigger value="opportunities" className="gap-2">
                <Target className="h-4 w-4" /> Oportunidades
              </TabsTrigger>
            </TabsList>

            {/* ===== OVERVIEW TAB ===== */}
            <TabsContent value="overview" className="space-y-3">
              {/* KPIs */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Linhas</span>
                    <div className="text-lg font-bold text-slate-800">{kpis.totalLinhas}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Quantidade</span>
                    <div className="text-lg font-bold text-slate-800">{kpis.totalQty.toLocaleString()}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Pedidos</span>
                    <div className="text-lg font-bold text-slate-800">{kpis.totalPedidos}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">OMs</span>
                    <div className="text-lg font-bold text-slate-800">{kpis.totalOMs}</div>
                  </CardContent>
                </Card>
              </div>

              {/* Filters - Collapsible */}
              <Collapsible open={showOverviewFilters} onOpenChange={setShowOverviewFilters}>
                <Card className="border shadow-sm bg-white">
                  <CollapsibleTrigger asChild>
                    <CardHeader className="pb-3 cursor-pointer select-none hover:bg-slate-50/50 transition-colors rounded-t-lg">
                      <div className="flex items-center justify-between flex-wrap gap-2">
                        <CardTitle className="text-lg flex items-center gap-2">
                          <Filter className="h-5 w-5" /> Filtros
                          {(empresaFilter.length + clienteFilter.length + equipamentoFilter.length + criticidadeFilter.length + statusFilter.length + periodFilter.length + estoqueFilter.length + (searchTerm ? 1 : 0)) > 0 && (
                            <Badge variant="secondary" className="ml-1 text-xs">{(empresaFilter.length + clienteFilter.length + equipamentoFilter.length + criticidadeFilter.length + statusFilter.length + periodFilter.length + estoqueFilter.length + (searchTerm ? 1 : 0))} ativos</Badge>
                          )}
                        </CardTitle>
                        <div className="flex items-center gap-2">
                          {(empresaFilter.length + clienteFilter.length + equipamentoFilter.length + criticidadeFilter.length + statusFilter.length + periodFilter.length + estoqueFilter.length + (searchTerm ? 1 : 0)) > 0 && (
                            <Button variant="ghost" size="sm" onClick={(e) => { e.stopPropagation(); clearFilters(); }}>
                              <X className="h-4 w-4 mr-1" /> Limpar
                            </Button>
                          )}
                          <ChevronDown className={`h-5 w-5 text-slate-400 transition-transform duration-200 ${showOverviewFilters ? 'rotate-180' : ''}`} />
                        </div>
                      </div>
                    </CardHeader>
                  </CollapsibleTrigger>
                  <CollapsibleContent>
                    <CardContent className="space-y-3">
                      <Input placeholder="Buscar por PN, descrição, cliente..." value={searchTerm} onChange={(e) => setSearchTerm(e.target.value)} />
                      <div className="flex flex-wrap gap-2">
                        <FilterDropdown label="Empresa" options={empresaOptions} selected={empresaFilter} onChange={setEmpresaFilter} />
                        <FilterDropdown label="Cliente" options={clienteOptions} selected={clienteFilter} onChange={setClienteFilter} />
                        <FilterDropdown label="Equipamento" options={equipamentoOptions} selected={equipamentoFilter} onChange={setEquipamentoFilter} />
                        <FilterDropdown label="Criticidade" options={criticidadeOptions} selected={criticidadeFilter} onChange={setCriticidadeFilter} />
                        <FilterDropdown label="Status" options={statusOptions} selected={statusFilter} onChange={setStatusFilter} />
                        <FilterDropdown label="Período" options={periodoOptions} selected={periodFilter} onChange={setPeriodFilter} />
                        <FilterDropdown label="Estoque" options={estoqueOptions} selected={estoqueFilter} onChange={setEstoqueFilter} />
                      </div>
                    </CardContent>
                  </CollapsibleContent>
                </Card>
              </Collapsible>

              {/* Charts - 2 columns */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-3">
                <Card className="border shadow-sm bg-white">
                  <CardHeader className="py-1"><CardTitle className="text-sm">Por Cliente (QTY)</CardTitle></CardHeader>
                  <CardContent>
                    <div className="h-[300px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={clientChartData} layout="vertical">
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis type="number" />
                          <YAxis dataKey="name" type="category" width={120} tick={{ fontSize: 10 }} />
                          <Tooltip />
                          <Legend />
                          <Bar dataKey="Quantidade" fill={COLORS.primary} />
                          <Bar dataKey="Vendido" fill={COLORS.success} />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>

                <Card className="border shadow-sm bg-white">
                  <CardHeader className="py-1"><CardTitle className="text-sm">Criticidade</CardTitle></CardHeader>
                  <CardContent>
                    <div className="h-[300px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie data={criticityChartData} cx="50%" cy="50%" labelLine={true} label={({ name, percent }) => `${name} (${(percent * 100).toFixed(0)}%)`} outerRadius={110} dataKey="value">
                            {criticityChartData.map((entry, index) => <Cell key={`cell-crit-${index}`} fill={entry.color} />)}
                          </Pie>
                          <Tooltip />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>

                <Card className="border shadow-sm bg-white">
                  <CardHeader className="py-1 flex-row flex-wrap items-center gap-2">
                    <CardTitle className="text-sm">Volume Mensal</CardTitle>
                    <div className="flex gap-2 ml-auto">
                      <Select value={chartYearFilter || '__all__'} onValueChange={(v) => setChartYearFilter(v === '__all__' ? '' : v)}>
                        <SelectTrigger className="w-[100px] h-7 text-xs"><SelectValue placeholder="Ano" /></SelectTrigger>
                        <SelectContent>
                          <SelectItem value="__all__">Todos</SelectItem>
                          {availableYears.map(y => (
                            <SelectItem key={y} value={y}>{y}</SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </div>
                  </CardHeader>
                  <CardContent>
                    <div className="h-[300px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <BarChart data={monthlyChartData}>
                          <CartesianGrid strokeDasharray="3 3" />
                          <XAxis dataKey="name" tick={{ fontSize: 10 }} />
                          <YAxis />
                          <Tooltip />
                          <Bar dataKey="value" fill={COLORS.primary} name="Quantidade" />
                        </BarChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>

                <Card className="border shadow-sm bg-white">
                  <CardHeader className="py-1 flex-row flex-wrap items-center gap-2">
                    <CardTitle className="text-sm">Ranking</CardTitle>
                    <div className="flex gap-1 ml-auto">
                      <Button variant={rankingType === 'inspecao' ? 'default' : 'outline'} size="sm" className="h-7 text-xs" onClick={() => setRankingType('inspecao')}>
                        Inspeção
                      </Button>
                      <Button variant={rankingType === 'venda' ? 'default' : 'outline'} size="sm" className="h-7 text-xs" onClick={() => setRankingType('venda')}>
                        Venda
                      </Button>
                    </div>
                  </CardHeader>
                  <CardContent>
                    <div className="max-h-[300px] overflow-y-auto">
                      <table className="w-full text-sm">
                        <thead className="sticky top-0 bg-white">
                          <tr className="border-b text-slate-500 text-xs">
                            <th className="text-left py-1 w-8">#</th>
                            <th className="text-left py-1">Equipamento</th>
                            <th className="text-right py-1 w-24">Inspeção</th>
                            <th className="text-right py-1 w-24">Venda</th>
                          </tr>
                        </thead>
                        <tbody>
                          {rankingData.map((item, i) => (
                            <tr key={item.name} className={`border-b border-slate-50 ${i === 0 ? 'bg-amber-50 font-semibold' : i === 1 ? 'bg-slate-50' : i === 2 ? 'bg-orange-50/50' : ''}`}>
                              <td className="py-1.5 text-slate-400">{i === 0 ? '🥇' : i === 1 ? '🥈' : i === 2 ? '🥉' : item.position}</td>
                              <td className="py-1.5 truncate max-w-[200px]" title={item.name}>{item.name}</td>
                              <td className={`py-1.5 text-right tabular-nums ${rankingType === 'inspecao' ? 'font-semibold text-slate-800' : 'text-slate-500'}`}>{item.inspecao.toLocaleString()}</td>
                              <td className={`py-1.5 text-right tabular-nums ${rankingType === 'venda' ? 'font-semibold text-slate-800' : 'text-slate-500'}`}>{item.venda.toLocaleString()}</td>
                            </tr>
                          ))}
                          {rankingData.length === 0 && (
                            <tr><td colSpan={4} className="py-8 text-center text-slate-400">Sem dados</td></tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </CardContent>
                </Card>

                <Card className="border shadow-sm bg-white">
                  <CardHeader className="py-1"><CardTitle className="text-sm">Distribuição por Origem</CardTitle></CardHeader>
                  <CardContent>
                    <div className="h-[300px]">
                      <ResponsiveContainer width="100%" height="100%">
                        <PieChart>
                          <Pie data={origemChartData} cx="50%" cy="50%" labelLine={true} label={({ name, percent }) => `${name} (${(percent * 100).toFixed(0)}%)`} outerRadius={100} dataKey="value">
                            {origemChartData.map((_, index) => (<Cell key={`cell-orig-${index}`} fill={PIE_COLORS[index % PIE_COLORS.length]} />))}
                          </Pie>
                          <Tooltip />
                        </PieChart>
                      </ResponsiveContainer>
                    </div>
                  </CardContent>
                </Card>
              </div>
            </TabsContent>

            {/* ===== OPPORTUNITIES TAB ===== */}
            <TabsContent value="opportunities" className="space-y-4">
              {/* Filters - Collapsible */}
              <Collapsible open={showOppFilters} onOpenChange={setShowOppFilters}>
                <Card className="border-0 shadow-md">
                  <CollapsibleTrigger asChild>
                    <CardHeader className="pb-3 cursor-pointer select-none hover:bg-slate-50/50 transition-colors rounded-t-lg">
                      <div className="flex items-center justify-between flex-wrap gap-2">
                        <CardTitle className="text-lg flex items-center gap-2">
                          <Filter className="h-5 w-5" /> Filtros
                          <Badge variant="outline">{oppFilteredData.length} de {data.length}</Badge>
                          {(oppMonthFilter.length + oppYearFilter.length + oppClienteFilter.length + oppStatusFilter.length + oppCriticidadeFilter.length + oppDiasFilter.length + oppAnaliseFilter.length + oppPrazoFilter.length + oppOrigemFilter.length + (oppSearchTerm ? 1 : 0)) > 0 && (
                            <Badge variant="secondary" className="text-xs">{(oppMonthFilter.length + oppYearFilter.length + oppClienteFilter.length + oppStatusFilter.length + oppCriticidadeFilter.length + oppDiasFilter.length + oppAnaliseFilter.length + oppPrazoFilter.length + oppOrigemFilter.length + (oppSearchTerm ? 1 : 0))} ativos</Badge>
                          )}
                        </CardTitle>
                        <div className="flex items-center gap-2">
                          {(oppMonthFilter.length + oppYearFilter.length + oppClienteFilter.length + oppStatusFilter.length + oppCriticidadeFilter.length + oppDiasFilter.length + oppAnaliseFilter.length + oppPrazoFilter.length + oppOrigemFilter.length + (oppSearchTerm ? 1 : 0)) > 0 && (
                            <Button variant="ghost" size="sm" onClick={(e) => { e.stopPropagation(); clearOppFilters(); }}>
                              <X className="h-4 w-4 mr-1" /> Limpar
                            </Button>
                          )}
                          <ChevronDown className={`h-5 w-5 text-slate-400 transition-transform duration-200 ${showOppFilters ? 'rotate-180' : ''}`} />
                        </div>
                      </div>
                    </CardHeader>
                  </CollapsibleTrigger>
                  <CollapsibleContent>
                    <CardContent className="space-y-3">
                      <Input placeholder="Buscar por PN, descrição, cliente..." value={oppSearchTerm} onChange={(e) => setOppSearchTerm(e.target.value)} />
                      <div className="flex flex-wrap gap-2">
                        <FilterDropdown label="Origem" options={origemOptions} selected={oppOrigemFilter} onChange={setOppOrigemFilter} />
                        <FilterDropdown label="Mês" options={monthOptions} selected={oppMonthFilter} onChange={setOppMonthFilter} />
                        <FilterDropdown label="Ano" options={yearOptions} selected={oppYearFilter} onChange={setOppYearFilter} />
                        <FilterDropdown label="Cliente" options={clienteOptions} selected={oppClienteFilter} onChange={setOppClienteFilter} />
                        <FilterDropdown label="Status" options={statusOptions} selected={oppStatusFilter} onChange={setOppStatusFilter} />
                        <FilterDropdown label="Criticidade" options={criticidadeOptions} selected={oppCriticidadeFilter} onChange={setOppCriticidadeFilter} />
                        <FilterDropdown label="Dias Aberto" options={diasOptions} selected={oppDiasFilter} onChange={setOppDiasFilter} />
                        <FilterDropdown label="Análise" options={analiseOptions} selected={oppAnaliseFilter} onChange={setOppAnaliseFilter} />
                        <FilterDropdown label="Prazo" options={prazoOptions} selected={oppPrazoFilter} onChange={setOppPrazoFilter} />
                      </div>
                    </CardContent>
                  </CollapsibleContent>
                </Card>
              </Collapsible>

              {/* KPIs */}
              <div className="grid grid-cols-3 sm:grid-cols-4 lg:grid-cols-7 gap-2">
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Linhas</span>
                    <div className="text-lg font-bold text-slate-800">{oppKpis.totalLinhas}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Qtd Total</span>
                    <div className="text-lg font-bold text-slate-800">{oppKpis.totalQty.toLocaleString()}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Pedidos</span>
                    <div className="text-lg font-bold text-slate-800">{oppKpis.totalPedidos}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">OMs</span>
                    <div className="text-lg font-bold text-slate-800">{oppKpis.totalOMs}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Filtradas</span>
                    <div className="text-lg font-bold text-slate-800">{oppKpis.linhasFiltradas}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Em Aberto</span>
                    <div className="text-lg font-bold text-slate-800">{oppKpis.totalEmAberto}</div>
                  </CardContent>
                </Card>
                <Card className="border shadow-sm bg-white">
                  <CardContent className="p-2">
                    <span className="text-[10px] uppercase tracking-wider text-slate-400 block">Total Vendido</span>
                    <div className="text-lg font-bold text-slate-800">{oppKpis.totalVendido.toLocaleString()}</div>
                    <span className="text-[10px] text-slate-400">{oppKpis.totalNF} NF · {oppKpis.faturadoParcial} parcial</span>
                  </CardContent>
                </Card>
              </div>

              {/* Data Table */}
              <Card className="border-0 shadow-md">
                <CardContent className="p-0">
                  <div className="overflow-auto" style={{ maxHeight: '675px' }}>
                    <Table>
                      <TableHeader className="sticky top-0 bg-slate-100 z-10">
                        <TableRow>
                          <TableHead className="font-semibold">Follow Up</TableHead>
                          <TableHead className="font-semibold">Equipamento</TableHead>
                          <TableHead className="font-semibold">PN</TableHead>
                          <TableHead className="font-semibold">Descrição</TableHead>
                          <TableHead className="font-semibold">Cliente</TableHead>
                          <TableHead className="font-semibold text-right">Qtd Inspeção</TableHead>
                          <TableHead className="font-semibold text-right">Qtd Pedida</TableHead>
                          <TableHead className="font-semibold text-right">Faturado</TableHead>
                          <TableHead className="font-semibold">Criticidade</TableHead>
                          <TableHead className="font-semibold">Status</TableHead>
                          <TableHead className="font-semibold">Data Abertura</TableHead>
                          <TableHead className="font-semibold text-right">Dias Aberto</TableHead>
                          <TableHead className="font-semibold">Pedido</TableHead>
                          <TableHead className="font-semibold">OM</TableHead>
                          <TableHead className="font-semibold text-right">LIC</TableHead>
                          <TableHead className="font-semibold text-right">Betim</TableHead>
                          <TableHead className="font-semibold text-right">Importação</TableHead>
                          <TableHead className="font-semibold text-center">Análise Estoque</TableHead>
                        </TableRow>
                      </TableHeader>
                      <TableBody>
                        {oppFilteredData.map((row) => (
                          <TableRow key={row.id} className={row.criticidade === 'Alta' ? 'bg-red-50' : ''}>
                            <TableCell>
                              <Button variant="ghost" size="sm" className="h-8 w-8 p-0 relative" onClick={() => openFollowUpModal(row)}>
                                {row.followUpLocal ? (
                                  <MessageCircle className="h-4 w-4 text-emerald-600" />
                                ) : row.followUpComercial ? (
                                  <span className="relative flex items-center justify-center">
                                    <MessageSquare className="h-4 w-4 text-amber-500" />
                                    <span className="absolute -top-1 -right-1 w-2.5 h-2.5 bg-amber-500 rounded-full border border-white" />
                                  </span>
                                ) : (
                                  <Plus className="h-4 w-4 text-slate-400" />
                                )}
                              </Button>
                            </TableCell>
                            <TableCell className="max-w-[160px] truncate" title={row.equipamento}>{row.equipamento || '-'}</TableCell>
                            <TableCell className="font-mono text-sm">{row.pn || '-'}</TableCell>
                            <TableCell className="max-w-[200px] truncate" title={row.partname || row.descricao}>{row.partname || row.descricao || '-'}</TableCell>
                            <TableCell>{row.cliente || '-'}</TableCell>
                            <TableCell className="text-right">{row.qty || 0}</TableCell>
                            <TableCell className="text-right">{row.quantidadePedida || 0}</TableCell>
                            <TableCell className="text-right font-medium">{row.quantidadeFaturada || 0}</TableCell>
                            <TableCell>
                              <Badge style={{ backgroundColor: CRITICITY_COLORS[row.criticidade] || '#6c757d', color: 'white' }}>
                                {row.criticidade || 'N/A'}
                              </Badge>
                            </TableCell>
                            <TableCell>
                              <Badge style={{ backgroundColor: STATUS_CONFIG[row.status]?.color || '#6c757d', color: 'white' }}>
                                {STATUS_CONFIG[row.status]?.label || row.status}
                              </Badge>
                              {row.status === 'vendido' && row.notaFiscal && (
                                <span className="ml-2 text-xs text-slate-500" title={`NF: ${row.notaFiscal}`}>NF: {row.notaFiscal}</span>
                              )}
                            </TableCell>
                            <TableCell>{formatDateBR(row.dataAbertura)}</TableCell>
                            <TableCell className="text-right font-medium" style={{ color: row.diasEmAberto > 60 ? '#dc3545' : 'inherit' }}>{row.diasEmAberto}</TableCell>
                            <TableCell className="font-mono text-sm">{row.pedidoCompra || '-'}</TableCell>
                            <TableCell className="font-mono text-sm">{row.ordemManutencao || '-'}</TableCell>
                            <TableCell className="text-right">{row.lic || 0}</TableCell>
                            <TableCell className="text-right">{row.betim || 0}</TableCell>
                            <TableCell className="text-right">{row.importacao || '-'}</TableCell>
                            <TableCell className="text-center">
                              {row.status === 'com_po_sem_faturamento' ? <AnaliseEstoqueBadge record={row} /> : <span className="text-slate-400">-</span>}
                            </TableCell>
                          </TableRow>
                        ))}
                      </TableBody>
                    </Table>
                  </div>
                </CardContent>
              </Card>
            </TabsContent>
          </Tabs>
        )}
      </main>

      {/* Footer */}
      <footer className="bg-white border-t mt-auto py-4 px-6 text-center text-sm text-slate-500">
        Dashboard de Oportunidades - Lundin Mining • Última atualização: {new Date().toLocaleDateString('pt-BR')} às {new Date().toLocaleTimeString('pt-BR')}
      </footer>
    </div>
  );
}
