import { useState, useEffect, useCallback, useRef } from 'react';
import { db } from './lib/firebase';
import { 
  collection, 
  query, 
  getDocs, 
  addDoc, 
  updateDoc, 
  deleteDoc, 
  doc, 
  writeBatch, 
  orderBy 
} from 'firebase/firestore';
import * as XLSX from 'xlsx';
import {
  X,
  Package,
  Building2,
  FileText,
  Filter,
  RefreshCw,
  ChevronDown,
  ChevronUp,
  Trash2,
  Edit3,
  Save,
  Loader2,
  AlertCircle,
  CheckCircle,
  Shield,
  Upload,
  FileSpreadsheet,
  Lock,
  Eye,
  EyeOff,
  Download,
  LogOut,
  Table,
} from 'lucide-react';

interface Ativo {
  id: string;
  placa: string;
  numero_loja: string;
  descricao: string;
  status: string;
  categoria: string;
  localizacao: string;
  data_aquisicao: string | null;
  valor: number;
  depr_acum: number;
  saldo_contabil: number;
  numero_bem: string;
  numero_incorporacao: string;
  observacao: string;
  created_at: string;
  updated_at: string;
}

type SortField = 'placa' | 'numero_loja' | 'descricao' | 'status' | 'categoria' | 'valor' | 'depr_acum' | 'saldo_contabil' | 'data_aquisicao' | 'created_at';
type SortDirection = 'asc' | 'desc';
type Page = 'main' | 'admin';

const ADMIN_PASSWORD = 'Oba2026Jatai';

const COLUMN_MAP: Record<string, string> = {
  placa: 'placa',
  'numero_loja': 'numero_loja',
  'numero da loja': 'numero_loja',
  'número da loja': 'numero_loja',
  'loja': 'numero_loja',
  'descricao': 'descricao',
  'descrição': 'descricao',
  'descrição do item': 'descricao',
  'descricao do item': 'descricao',
  'item': 'descricao',
  'status': 'status',
  'categoria': 'categoria',
  'localizacao': 'localizacao',
  'localização': 'localizacao',
  'local': 'localizacao',
  'data_aquisicao': 'data_aquisicao',
  'data aquisição': 'data_aquisicao',
  'data aquisicao': 'data_aquisicao',
  'data de aquisição': 'data_aquisicao',
  'data de aquisicao': 'data_aquisicao',
  'aquisicao': 'data_aquisicao',
  'aquisição': 'data_aquisicao',
  'valor': 'valor',
  'depr_acum': 'depr_acum',
  'depr. acum': 'depr_acum',
  'depreciação acumulada': 'depr_acum',
  'depreciacao acumulada': 'depr_acum',
  'depr. acumulada': 'depr_acum',
  'depr acum': 'depr_acum',
  'saldo_contabil': 'saldo_contabil',
  'saldo contábil': 'saldo_contabil',
  'saldo_contábil': 'saldo_contabil',
  'numero_bem': 'numero_bem',
  'numero do bem': 'numero_bem',
  'número do bem': 'numero_bem',
  'nº bem': 'numero_bem',
  'numero_incorporacao': 'numero_incorporacao',
  'numero da incorporação': 'numero_incorporacao',
  'número da incorporação': 'numero_incorporacao',
  'nº incorporação': 'numero_incorporacao',
  'observacao': 'observacao',
  'observação': 'observacao',
  'obs': 'observacao',
  'notas': 'observacao',
};

function mapExcelRow(row: Record<string, unknown>): Partial<Ativo> {
  const mapped: Record<string, unknown> = {};
  for (const [header, value] of Object.entries(row)) {
    const key = header.toLowerCase().trim();
    const dbField = COLUMN_MAP[key];
    if (dbField) {
      mapped[dbField] = value;
    }
  }
  return mapped as Partial<Ativo>;
}

const parseNumericValue = (val: any): number => {
  if (typeof val === 'number') return val;
  if (!val || typeof val !== 'string') return 0;
  // Remove pontos de milhar e substitui a vírgula decimal por ponto
  const sanitized = val.trim().replace(/\./g, '').replace(',', '.');
  const num = parseFloat(sanitized);
  return isNaN(num) ? 0 : num;
};

const dbFields = ['placa', 'numero_loja', 'numero_bem', 'numero_incorporacao', 'descricao', 'status', 'categoria', 'localizacao', 'data_aquisicao', 'valor', 'depr_acum', 'saldo_contabil', 'observacao'];
const fieldLabels: Record<string, string> = {
  placa: 'Placa',
  numero_loja: 'Número da Loja',
  descricao: 'Descrição',
  status: 'Status',
  categoria: 'Categoria',
  localizacao: 'Localização',
  data_aquisicao: 'Data Aquisição',
  valor: 'Valor',
  depr_acum: 'Depr. Acum',
  saldo_contabil: 'Saldo Contábil',
  numero_bem: 'Número do Bem',
  numero_incorporacao: 'Incorporação',
  observacao: 'Observação',
};

function App() {
  const [page, setPage] = useState<Page>('main');
  const [isAdminAuth, setIsAdminAuth] = useState(false);
  const [passwordInput, setPasswordInput] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [passwordError, setPasswordError] = useState('');

  // Excel import states
  const [excelRawData, setExcelRawData] = useState<Record<string, unknown>[]>([]);
  const [excelColumns, setExcelColumns] = useState<string[]>([]);
  const [excelData, setExcelData] = useState<Partial<Ativo>[]>([]);
  const [excelFileName, setExcelFileName] = useState('');
  const [importing, setImporting] = useState(false);
  const [importResult, setImportResult] = useState<{ success: number; errors: number } | null>(null);
  const [columnMapping, setColumnMapping] = useState<Record<string, string>>({});
  const [showMappingStep, setShowMappingStep] = useState(false);
  const [visibleColumns, setVisibleColumns] = useState<Set<string>>(new Set(dbFields));
  const fileInputRef = useRef<HTMLInputElement>(null);

  // Main page states
  const [ativos, setAtivos] = useState<Ativo[]>([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [successMessage, setSuccessMessage] = useState<string | null>(null);

  const [filtroPlaca, setFiltroPlaca] = useState('');
  const [filtroLoja, setFiltroLoja] = useState('');
  const [filtroDescricao, setFiltroDescricao] = useState('');
  const [filtroNumeroBem, setFiltroNumeroBem] = useState('');
  const [filtroNumeroIncorporacao, setFiltroNumeroIncorporacao] = useState('');
  const [displayLimit, setDisplayLimit] = useState(10); // Novo estado para o limite de exibição

  const [sortField, setSortField] = useState<SortField>('created_at');
  const [sortDirection, setSortDirection] = useState<SortDirection>('desc');

  const [showAddModal, setShowAddModal] = useState(false);
  const [newAtivo, setNewAtivo] = useState({
    placa: '', numero_loja: '', descricao: '', status: 'ativo',
    categoria: '', localizacao: '', data_aquisicao: '', valor: '',
    depr_acum: '', saldo_contabil: '', numero_bem: '', numero_incorporacao: '', observacao: '',
  });

  const [selectedAtivo, setSelectedAtivo] = useState<Ativo | null>(null);
  const [editingAtivo, setEditingAtivo] = useState<Ativo | null>(null);
  const [editForm, setEditForm] = useState({
    placa: '', numero_loja: '', descricao: '', status: 'ativo',
    categoria: '', localizacao: '', data_aquisicao: '', valor: '', depr_acum: '', saldo_contabil: '',
    numero_bem: '', numero_incorporacao: '', observacao: '',
  });

  const [deletingId, setDeletingId] = useState<string | null>(null);

  const fetchAtivos = useCallback(async () => {
    setLoading(true);
    setError(null);
    setDisplayLimit(10); // Reseta o limite ao buscar ou filtrar
    try {
      const ativosRef = collection(db, 'ativos');
      const q = query(ativosRef, orderBy('created_at', 'desc'));
      const querySnapshot = await getDocs(q);
      
      const allAtivos = querySnapshot.docs.map(doc => ({
        id: doc.id,
        ...doc.data()
      })) as Ativo[];

      // Filtra em memória para suportar busca parcial case-insensitive (substituindo ilike)
      const filtered = allAtivos.filter(ativo => {
        const matchPlaca = !filtroPlaca.trim() || (ativo.placa || '').includes(filtroPlaca.trim());
        const matchLoja = !filtroLoja.trim() || ativo.numero_loja === filtroLoja.trim();
        const matchDesc = !filtroDescricao.trim() || (ativo.descricao || '').toLowerCase().includes(filtroDescricao.trim().toLowerCase());
        const matchBem = !filtroNumeroBem.trim() || (ativo.numero_bem || '').includes(filtroNumeroBem.trim());
        const matchInc = !filtroNumeroIncorporacao.trim() || (ativo.numero_incorporacao || '').includes(filtroNumeroIncorporacao.trim());
        return matchPlaca && matchLoja && matchDesc && matchBem && matchInc;
      });

      setAtivos(filtered);
    } catch (err) {
      setError(err instanceof Error ? err.message : 'Erro ao carregar ativos');
    } finally {
      setLoading(false);
    }
  }, [filtroPlaca, filtroLoja, filtroDescricao, filtroNumeroBem, filtroNumeroIncorporacao]);

  useEffect(() => { fetchAtivos(); }, [fetchAtivos]);
  useEffect(() => {
    if (successMessage) { const t = setTimeout(() => setSuccessMessage(null), 3000); return () => clearTimeout(t); }
  }, [successMessage]);

  const handleSort = (field: SortField) => {
    if (sortField === field) setSortDirection(sortDirection === 'asc' ? 'desc' : 'asc');
    else { setSortField(field); setSortDirection('asc'); }
  };

  const sortedAtivos = [...ativos].sort((a, b) => {
    const dir = sortDirection === 'asc' ? 1 : -1;
    const aVal = a[sortField] ?? '';
    const bVal = b[sortField] ?? '';
    if (sortField === 'valor' || sortField === 'depr_acum' || sortField === 'saldo_contabil') return dir * ((a[sortField as keyof Ativo] as number ?? 0) - (b[sortField as keyof Ativo] as number ?? 0));
    return dir * String(aVal).localeCompare(String(bVal), 'pt-BR');
  });

  const displayedAtivos = sortedAtivos.slice(0, displayLimit);

  const handleAddAtivo = async () => {
    try {
      await addDoc(collection(db, 'ativos'), {
        placa: newAtivo.placa, 
        numero_loja: newAtivo.numero_loja, 
        descricao: newAtivo.descricao,
        status: newAtivo.status, 
        categoria: newAtivo.categoria, 
        localizacao: newAtivo.localizacao,
        data_aquisicao: newAtivo.data_aquisicao || null,
        valor: parseNumericValue(newAtivo.valor),
        depr_acum: parseNumericValue(newAtivo.depr_acum),
        saldo_contabil: parseNumericValue(newAtivo.saldo_contabil),
        numero_bem: newAtivo.numero_bem,
        numero_incorporacao: newAtivo.numero_incorporacao,
        observacao: newAtivo.observacao,
        created_at: new Date().toISOString(),
        updated_at: new Date().toISOString()
      });
      setShowAddModal(false);
      setNewAtivo({ placa: '', numero_loja: '', descricao: '', status: 'ativo', categoria: '', localizacao: '', data_aquisicao: '', valor: '', depr_acum: '', saldo_contabil: '', numero_bem: '', numero_incorporacao: '', observacao: '' });
      setSuccessMessage('Ativo adicionado com sucesso!');
      fetchAtivos();
    } catch (err) { setError(err instanceof Error ? err.message : 'Erro ao adicionar ativo'); }
  };

  const handleEditAtivo = async () => {
    if (!editingAtivo) return;
    try {
      await updateDoc(doc(db, 'ativos', editingAtivo.id), {
        placa: editForm.placa, 
        numero_loja: editForm.numero_loja, 
        descricao: editForm.descricao,
        status: editForm.status, 
        categoria: editForm.categoria, 
        localizacao: editForm.localizacao,
        data_aquisicao: editForm.data_aquisicao || null,
        valor: parseNumericValue(editForm.valor),
        depr_acum: parseNumericValue(editForm.depr_acum),
        saldo_contabil: parseNumericValue(editForm.saldo_contabil),
        numero_bem: editForm.numero_bem,
        numero_incorporacao: editForm.numero_incorporacao,
        observacao: editForm.observacao, 
        updated_at: new Date().toISOString(),
      });
      setEditingAtivo(null);
      setSuccessMessage('Ativo atualizado com sucesso!');
      fetchAtivos();
    } catch (err) { setError(err instanceof Error ? err.message : 'Erro ao atualizar ativo'); }
  };

  const handleDeleteAtivo = async (id: string) => {
    try {
      await deleteDoc(doc(db, 'ativos', id));
      setDeletingId(null);
      setSuccessMessage('Ativo removido com sucesso!');
      fetchAtivos();
    } catch (err) { setError(err instanceof Error ? err.message : 'Erro ao remover ativo'); }
  };

  const openEditModal = (ativo: Ativo) => {
    setEditingAtivo(ativo);
    setEditForm({
      placa: ativo.placa, numero_loja: ativo.numero_loja, descricao: ativo.descricao,
      status: ativo.status, categoria: ativo.categoria || '', localizacao: ativo.localizacao || '',
      data_aquisicao: ativo.data_aquisicao || '', valor: ativo.valor?.toString() || '',
      depr_acum: ativo.depr_acum?.toString() || '', saldo_contabil: ativo.saldo_contabil?.toString() || '',
      numero_bem: ativo.numero_bem || '', numero_incorporacao: ativo.numero_incorporacao || '',
      observacao: ativo.observacao || '',
    });
  };

  const handleAdminLogin = () => {
    if (passwordInput === ADMIN_PASSWORD) {
      setIsAdminAuth(true);
      setPasswordError('');
      setPasswordInput('');
    } else {
      setPasswordError('Senha incorreta');
    }
  };

  const handleExcelUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    setExcelFileName(file.name);
    setImportResult(null);
    setColumnMapping({});
    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target?.result;
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });
      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet);
      setExcelRawData(jsonData);
      const cols = jsonData.length > 0 ? Object.keys(jsonData[0]) : [];
      setExcelColumns(cols);

      // Pré-preenchimento automático do mapeamento baseando-se no COLUMN_MAP
      const initialMapping: Record<string, string> = {};
      dbFields.forEach(field => {
        const foundCol = cols.find(col => {
          const key = col.toLowerCase().trim();
          return COLUMN_MAP[key] === field;
        });
        if (foundCol) initialMapping[field] = foundCol;
      });
      setColumnMapping(initialMapping);

      setShowMappingStep(true);
      setExcelData([]);
    };
    reader.readAsArrayBuffer(file);
  };

  const applyCustomMapping = () => {
    const mapped = excelRawData.map(row => {
      const mappedRow: Record<string, unknown> = {};
      Object.entries(columnMapping).forEach(([dbField, excelCol]) => {
        if (excelCol) mappedRow[dbField] = row[excelCol];
      });
      return mappedRow as Partial<Ativo>;
    }).filter(row => row.placa || row.descricao || row.numero_loja);
    setExcelData(mapped);
    setShowMappingStep(false);
  };

  const toggleColumnVisibility = (field: string) => {
    const newVisible = new Set(visibleColumns);
    if (newVisible.has(field)) {
      newVisible.delete(field);
    } else {
      newVisible.add(field);
    }
    setVisibleColumns(newVisible);
  };

  const handleImportToDatabase = async () => {
    if (excelData.length === 0) return;
    setImporting(true);
    setImportResult(null);
    let successCount = 0;
    let errorCount = 0;
    const BATCH_SIZE = 50;

    const formatToPostgresDate = (val: any): string | null => {
      if (!val) return null;
      const d = new Date(val);
      if (isNaN(d.getTime())) return null;
      return d.toISOString().split('T')[0];
    };

    for (let i = 0; i < excelData.length; i += BATCH_SIZE) {
      try {
        const batch = writeBatch(db);
        const chunk = excelData.slice(i, i + BATCH_SIZE);
        
        chunk.forEach(row => {
          const docRef = doc(collection(db, 'ativos'));
          batch.set(docRef, {
            placa: String(row.placa || ''),
            numero_loja: String(row.numero_loja || ''),
            descricao: String(row.descricao || ''),
            status: String(row.status || 'ativo'),
            categoria: String(row.categoria || ''),
            localizacao: String(row.localizacao || ''),
            data_aquisicao: formatToPostgresDate(row.data_aquisicao),
            valor: parseNumericValue(row.valor),
            depr_acum: parseNumericValue(row.depr_acum),
            saldo_contabil: parseNumericValue(row.saldo_contabil),
            numero_bem: String(row.numero_bem || ''),
            numero_incorporacao: String(row.numero_incorporacao || ''),
            observacao: String(row.observacao || ''),
            created_at: new Date().toISOString(),
            updated_at: new Date().toISOString()
          });
        });
        
        await batch.commit();
        successCount += chunk.length;
      } catch (err) {
        console.error('Erro ao inserir lote no Firebase:', err);
        errorCount += chunk.length;
      }
    }
    setImportResult({ success: successCount, errors: errorCount });
    setImporting(false);
    fetchAtivos();
  };

  const handleExportExcel = () => {
    const dataToExport = sortedAtivos.length > 0 ? sortedAtivos : ativos;
    if (dataToExport.length === 0) return;

    const exportData = dataToExport.map(a => {
      const row: Record<string, any> = {};
      dbFields.forEach(field => {
        const label = fieldLabels[field];
        row[label] = a[field as keyof Ativo] ?? '';
      });
      return row;
    });

    const ws = XLSX.utils.json_to_sheet(exportData);
    
    // Auto-ajuste de colunas no Excel
    const colWidths = Object.keys(exportData[0] || {}).map(key => {
      const maxLen = Math.max(
        key.length,
        ...exportData.slice(0, 1000).map(row => String(row[key] || '').length)
      );
      return { wch: maxLen + 2 };
    });
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Ativos');
    XLSX.writeFile(wb, `ativos_export_${new Date().toISOString().split('T')[0]}.xlsx`);
  };

  const handleDownloadTemplate = () => {
    const templateRow: Record<string, any> = {};
    // Cria cabeçalhos baseados nos labels amigáveis
    dbFields.forEach(field => {
      templateRow[fieldLabels[field]] = '';
    });
    
    // Adiciona uma linha de exemplo
    templateRow['Placa'] = 'ABC-1234';
    templateRow['Número da Loja'] = '01';
    templateRow['Descrição'] = 'Exemplo de Item';
    templateRow['Status'] = 'ativo';
    
    const ws = XLSX.utils.json_to_sheet([templateRow]);

    // Auto-ajuste de colunas para o Template
    const colWidths = Object.keys(templateRow).map(key => ({
      wch: Math.max(key.length, String(templateRow[key]).length) + 5
    }));
    ws['!cols'] = colWidths;

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Template');
    XLSX.writeFile(wb, 'modelo_importacao_ativos.xlsx');
  };

  const handleClearDatabase = async () => {
    if (!confirm('ATENCAO: Isso ira apagar TODOS os ativos do banco de dados. Deseja continuar?')) return;
    setLoading(true);
    try {
      const querySnapshot = await getDocs(collection(db, 'ativos'));
      const batchSize = 500; // Limite do Firestore por lote
      let batch = writeBatch(db);
      let count = 0;

      for (const snapshot of querySnapshot.docs) {
        batch.delete(snapshot.ref);
        count++;
        if (count === batchSize) {
          await batch.commit();
          batch = writeBatch(db);
          count = 0;
        }
      }
      if (count > 0) await batch.commit();

      setSuccessMessage('Todos os ativos foram removidos do banco de dados.');
      fetchAtivos();
    } catch (err) { 
      setError(err instanceof Error ? err.message : 'Erro ao limpar banco de dados'); 
    } finally {
      setLoading(false);
    }
  };

  const SortIcon = ({ field }: { field: SortField }) => {
    if (sortField !== field) return <ChevronDown className="w-3.5 h-3.5 text-slate-400" />;
    return sortDirection === 'asc' ? <ChevronUp className="w-3.5 h-3.5 text-emerald-600" /> : <ChevronDown className="w-3.5 h-3.5 text-emerald-600" />;
  };

  const formatCurrency = (value: number) => new Intl.NumberFormat('pt-BR', { style: 'currency', currency: 'BRL' }).format(value);
  const formatDate = (date: string | null) => { if (!date) return '-'; return new Intl.DateTimeFormat('pt-BR').format(new Date(date)); };

  const statusColors: Record<string, string> = {
    ativo: 'bg-emerald-50 text-emerald-700 border-emerald-200',
    inativo: 'bg-red-50 text-red-700 border-red-200',
    manutencao: 'bg-amber-50 text-amber-700 border-amber-200',
  };

  // ===================== ADMIN PAGE =====================
  if (page === 'admin') {
    if (!isAdminAuth) {
      return (
        <div className="min-h-screen bg-slate-50 flex items-center justify-center">
          <div className="w-full max-w-md mx-4">
            <div className="bg-white rounded-2xl shadow-xl border border-slate-200 overflow-hidden">
              <div className="p-8 text-center border-b border-slate-100">
                <div className="w-16 h-16 mx-auto mb-4 bg-slate-900 rounded-2xl flex items-center justify-center">
                  <Shield className="w-8 h-8 text-white" />
                </div>
                <h2 className="text-xl font-bold text-slate-900">Area Administrativa</h2>
                <p className="text-sm text-slate-500 mt-1">Insira a senha para acessar</p>
              </div>
              <div className="p-6">
                {passwordError && (
                  <div className="mb-4 flex items-center gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">
                    <AlertCircle className="w-4 h-4 flex-shrink-0" />
                    {passwordError}
                  </div>
                )}
                <div className="space-y-4">
                  <div>
                    <label className="flex items-center gap-1.5 text-xs font-medium text-slate-500 mb-1.5">
                      <Lock className="w-3.5 h-3.5" />
                      Senha de acesso
                    </label>
                    <div className="relative">
                      <input
                        type={showPassword ? 'text' : 'password'}
                        value={passwordInput}
                        onChange={(e) => { setPasswordInput(e.target.value); setPasswordError(''); }}
                        onKeyDown={(e) => e.key === 'Enter' && handleAdminLogin()}
                        placeholder="Digite a senha..."
                        className="w-full pl-3 pr-10 py-3 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-slate-500/20 focus:border-slate-500"
                      />
                      <button
                        onClick={() => setShowPassword(!showPassword)}
                        className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"
                      >
                        {showPassword ? <EyeOff className="w-4 h-4" /> : <Eye className="w-4 h-4" />}
                      </button>
                    </div>
                  </div>
                  <button
                    onClick={handleAdminLogin}
                    className="w-full py-3 text-sm font-semibold text-white bg-slate-900 rounded-lg hover:bg-slate-800 transition-colors"
                  >
                    Entrar
                  </button>
                  <button
                    onClick={() => { setPage('main'); setPasswordError(''); setPasswordInput(''); }}
                    className="w-full py-2.5 text-sm font-medium text-slate-500 bg-slate-50 border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors"
                  >
                    Voltar para consulta
                  </button>
                </div>
              </div>
            </div>
          </div>
        </div>
      );
    }

    // Admin authenticated view
    return (
      <div className="min-h-screen bg-slate-50">
        <header className="bg-slate-900 sticky top-0 z-30">
          <div className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8">
            <div className="flex items-center justify-between h-16">
              <div className="flex items-center gap-3">
                <div className="w-10 h-10 bg-white/10 rounded-xl flex items-center justify-center">
                  <Shield className="w-5 h-5 text-white" />
                </div>
                <div>
                  <h1 className="text-lg font-semibold text-white leading-tight">Admin - Importar Planilha</h1>
                  <p className="text-xs text-slate-400">Carregar dados do Excel para o banco de dados</p>
                </div>
              </div>
              <button
                onClick={() => { setPage('main'); }}
                className="inline-flex items-center gap-2 px-4 py-2.5 bg-white/10 text-white text-sm font-medium rounded-lg hover:bg-white/20 transition-all"
              >
                <LogOut className="w-4 h-4" />
                Voltar
              </button>
            </div>
          </div>
        </header>

        <main className="max-w-5xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
          {successMessage && (
            <div className="mb-6 flex items-center gap-2 p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-700 text-sm">
              <CheckCircle className="w-4 h-4 flex-shrink-0" />
              {successMessage}
            </div>
          )}
          {error && (
            <div className="mb-6 flex items-center gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">
              <AlertCircle className="w-4 h-4 flex-shrink-0" />
              {error}
              <button onClick={() => setError(null)} className="ml-auto"><X className="w-4 h-4" /></button>
            </div>
          )}

          {/* Upload Section */}
          <div className="bg-white rounded-xl border border-slate-200 shadow-sm mb-6">
            <div className="px-6 py-4 border-b border-slate-100">
              <div className="flex items-center gap-2 text-sm font-semibold text-slate-800">
                <Upload className="w-4 h-4 text-slate-600" />
                Importar Planilha Excel
              </div>
            </div>
            <div className="p-6">
              <div
                onClick={() => fileInputRef.current?.click()}
                className="border-2 border-dashed border-slate-200 rounded-xl p-8 text-center cursor-pointer hover:border-slate-400 hover:bg-slate-50/50 transition-all group"
              >
                <div className="w-14 h-14 mx-auto mb-4 bg-slate-100 rounded-2xl flex items-center justify-center group-hover:bg-slate-200 transition-colors">
                  <FileSpreadsheet className="w-7 h-7 text-slate-500" />
                </div>
                <p className="text-sm font-medium text-slate-700">
                  Clique para selecionar uma planilha
                </p>
                <p className="text-xs text-slate-400 mt-1">
                  Formatos aceitos: .xlsx, .xls, .csv
                </p>
                <div className="mt-4 flex justify-center">
                  <button
                    onClick={(e) => { e.stopPropagation(); handleDownloadTemplate(); }}
                    className="inline-flex items-center gap-2 px-3 py-1.5 text-xs font-medium text-emerald-700 bg-emerald-50 border border-emerald-200 rounded-lg hover:bg-emerald-100 transition-colors"
                  >
                    <Download className="w-3.5 h-3.5" />
                    Baixar Planilha Modelo
                  </button>
                </div>
                {excelFileName && (
                  <div className="mt-3 inline-flex items-center gap-2 px-3 py-1.5 bg-emerald-50 border border-emerald-200 rounded-lg text-sm text-emerald-700">
                    <FileSpreadsheet className="w-4 h-4" />
                    {excelFileName}
                  </div>
                )}
              </div>
              <input
                ref={fileInputRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                onChange={handleExcelUpload}
                className="hidden"
              />

              {/* Column mapping setup */}
              {showMappingStep && excelColumns.length > 0 && (
                <div className="mt-6 p-4 bg-blue-50 rounded-lg border border-blue-200">
                  <h4 className="text-xs font-semibold text-blue-900 mb-3 flex items-center gap-1.5">
                    <Table className="w-3.5 h-3.5" />
                    Mapeamento de colunas da planilha
                  </h4>
                  <p className="text-xs text-blue-800 mb-4">
                    Selecione qual coluna da planilha corresponde a cada campo do sistema:
                  </p>
                  <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 mb-4">
                    {dbFields.map(field => (
                      <div key={field}>
                        <label className="text-xs font-medium text-slate-700 mb-1.5 block capitalize">
                          {fieldLabels[field] || field}
                        </label>
                        <select
                          value={columnMapping[field] || ''}
                          onChange={(e) => setColumnMapping({ ...columnMapping, [field]: e.target.value })}
                          className="w-full px-2 py-1.5 bg-white border border-slate-300 rounded text-xs focus:outline-none focus:ring-2 focus:ring-blue-500/20 focus:border-blue-500"
                        >
                          <option value="">-- Não mapear --</option>
                          {excelColumns.map(col => (
                            <option key={col} value={col}>{col}</option>
                          ))}
                        </select>
                      </div>
                    ))}
                  </div>
                  <button
                    onClick={applyCustomMapping}
                    className="w-full py-2 text-sm font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 transition-colors"
                  >
                    Aplicar Mapeamento
                  </button>
                </div>
              )}

              {/* Column mapping info - shown when NOT in mapping step */}
              {!showMappingStep && excelColumns.length === 0 && (
                <div className="mt-6 p-4 bg-slate-50 rounded-lg border border-slate-200">
                  <h4 className="text-xs font-semibold text-slate-600 mb-2 flex items-center gap-1.5">
                    <Table className="w-3.5 h-3.5" />
                    Mapeamento automatico de colunas
                  </h4>
                  <p className="text-xs text-slate-500 mb-3">
                    A planilha sera lida automaticamente. Use os seguintes nomes de coluna para mapeamento:
                  </p>
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                    {[
                      ['placa', 'Placa do ativo'],
                      ['numero_loja / loja', 'Numero da loja'],
                      ['numero_bem', 'Número do bem'],
                      ['numero_incorporacao', 'Número da incorporação'],
                      ['descricao / item', 'Descricao do item'],
                      ['status', 'Status (ativo/inativo)'],
                      ['categoria', 'Categoria'],
                      ['localizacao / local', 'Localizacao'],
                      ['data_aquisicao', 'Data de aquisicao'],
                      ['valor', 'Valor em R$'],
                      ['depr_acum', 'Depreciacao acumulada'],
                      ['saldo_contabil', 'Saldo contabil'],
                      ['observacao / obs', 'Observacao'],
                    ].map(([col, desc]) => (
                      <div key={col} className="flex items-start gap-1.5">
                        <code className="text-xs bg-white px-1.5 py-0.5 rounded border border-slate-200 text-slate-700 font-mono">{col}</code>
                        <span className="text-xs text-slate-400">{desc}</span>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
            </div>

          {/* Spreadsheet Viewer and Column Selector */}
          {excelRawData.length > 0 && !showMappingStep && (
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm mb-6">
              <div className="px-6 py-4 border-b border-slate-100">
                <h3 className="text-sm font-semibold text-slate-800 flex items-center gap-2">
                  <FileSpreadsheet className="w-4 h-4" />
                  Visualizar Planilha
                </h3>
              </div>
              <div className="p-6 space-y-4">
                {/* Column Visibility Selector */}
                <div>
                  <label className="text-xs font-semibold text-slate-600 mb-3 block">
                    Colunas visíveis na tabela de ativos:
                  </label>
                  <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-4 gap-2">
                    {dbFields.map(field => (
                      <button
                        key={field}
                        onClick={() => toggleColumnVisibility(field)}
                        className={`flex items-center gap-2 px-3 py-2 rounded-lg border text-sm transition-all ${
                          visibleColumns.has(field)
                            ? 'bg-emerald-50 border-emerald-300 text-emerald-700 font-medium'
                            : 'bg-slate-50 border-slate-300 text-slate-600 hover:border-slate-400'
                        }`}
                      >
                        <div className={`w-4 h-4 rounded border-2 flex items-center justify-center ${
                          visibleColumns.has(field) ? 'bg-emerald-600 border-emerald-600' : 'border-slate-300'
                        }`}>
                          {visibleColumns.has(field) && <span className="text-white text-xs">✓</span>}
                        </div>
                        {fieldLabels[field]}
                      </button>
                    ))}
                  </div>
                </div>

                {/* Spreadsheet Preview */}
                <div>
                  <label className="text-xs font-semibold text-slate-600 mb-2 block">
                    Dados da planilha ({excelRawData.length} linhas):
                  </label>
                  <div className="overflow-x-auto border border-slate-200 rounded-lg">
                    <table className="w-full text-xs">
                      <thead className="bg-slate-100 sticky top-0">
                        <tr>
                          <th className="px-3 py-2 text-left font-semibold text-slate-700 w-12">#</th>
                          {excelColumns.map(col => (
                            <th key={col} className="px-3 py-2 text-left font-semibold text-slate-700 whitespace-nowrap">
                              {col}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-slate-100">
                        {excelRawData.slice(0, 50).map((row, i) => (
                          <tr key={i} className="hover:bg-slate-50">
                            <td className="px-3 py-2 text-slate-400 bg-slate-50 font-mono">{i + 1}</td>
                            {excelColumns.map(col => (
                              <td key={col} className="px-3 py-2 text-slate-700 max-w-xs truncate">
                                {String(row[col] || '-')}
                              </td>
                            ))}
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    {excelRawData.length > 50 && (
                      <div className="px-3 py-2 text-center text-xs text-slate-400 bg-slate-50 border-t">
                        Mostrando 50 de {excelRawData.length} registros
                      </div>
                    )}
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* Preview Section */}

          {excelData.length > 0 && (
            <div className="bg-white rounded-xl border border-slate-200 shadow-sm mb-6">
              <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <h3 className="text-sm font-semibold text-slate-800">Previa dos Dados</h3>
                  <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-emerald-50 text-emerald-700 border border-emerald-200">
                    {excelData.length} registros
                  </span>
                </div>
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => { setExcelData([]); setExcelRawData([]); setExcelColumns([]); setExcelFileName(''); setImportResult(null); setColumnMapping({}); setShowMappingStep(false); }}
                    className="px-3 py-1.5 text-xs font-medium text-slate-600 bg-slate-50 border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors"
                  >
                    Limpar
                  </button>
                  <button
                    onClick={handleImportToDatabase}
                    disabled={importing}
                    className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-white bg-emerald-600 rounded-lg hover:bg-emerald-700 transition-colors disabled:opacity-50"
                  >
                    {importing ? <Loader2 className="w-4 h-4 animate-spin" /> : <Upload className="w-4 h-4" />}
                    {importing ? 'Importando...' : 'Importar para o Banco'}
                  </button>
                </div>
              </div>
              <div className="overflow-x-auto max-h-80">
                <table className="w-full text-xs">
                  <thead className="sticky top-0 bg-slate-50">
                    <tr>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">#</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap w-8">Placa</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap w-10">Loja</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Descrição</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Status</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Categoria</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Localização</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Aquisição</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Valor</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Depr. Acum</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Saldo Contábil</th>
                      <th className="px-2 py-2.5 text-left text-xs font-semibold text-slate-500 whitespace-nowrap">Obs</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {excelData.slice(0, 100).map((row, i) => (
                      <tr key={i} className="hover:bg-slate-50/50">
                        <td className="px-2 py-2 text-slate-400 text-xs">{i + 1}</td>
                        <td className="px-2 py-2 font-mono text-slate-700 whitespace-nowrap w-8">{String(row.placa || '-')}</td>
                        <td className="px-2 py-2 text-slate-700 whitespace-nowrap w-10">{String(row.numero_loja || '-')}</td>
                        <td className="px-2 py-2 text-slate-700 max-w-md truncate">{String(row.descricao || '-')}</td>
                        <td className="px-2 py-2">
                          <span className={`inline-flex px-2 py-0.5 text-xs font-medium rounded-full border whitespace-nowrap ${statusColors[String(row.status || 'ativo')] || 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                            {String(row.status || 'ativo')}
                          </span>
                        </td>
                        <td className="px-2 py-2 text-slate-700 whitespace-nowrap">{String(row.categoria || '-')}</td>
                        <td className="px-2 py-2 text-slate-700 whitespace-nowrap">{String(row.localizacao || '-')}</td>
                        <td className="px-2 py-2 text-slate-700 whitespace-nowrap">{row.data_aquisicao ? formatDate(String(row.data_aquisicao)) : '-'}</td>
                        <td className="px-2 py-2 font-mono text-slate-700 whitespace-nowrap">{row.valor ? formatCurrency(parseNumericValue(row.valor)) : '-'}</td>
                        <td className="px-2 py-2 font-mono text-slate-700 whitespace-nowrap">{row.depr_acum ? formatCurrency(parseNumericValue(row.depr_acum)) : '-'}</td>
                        <td className="px-2 py-2 font-mono text-slate-700 whitespace-nowrap">{row.saldo_contabil ? formatCurrency(parseNumericValue(row.saldo_contabil)) : '-'}</td>
                        <td className="px-2 py-2 text-slate-700 max-w-xs truncate">{String(row.observacao || '-')}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
                {excelData.length > 100 && (
                  <div className="px-4 py-3 text-center text-xs text-slate-400 border-t border-slate-100">
                    Mostrando 100 de {excelData.length} registros
                  </div>
                )}
              </div>
            </div>
          )}

          {/* Import Result */}
          {importResult && (
            <div className={`mb-6 rounded-xl border p-4 ${importResult.errors > 0 ? 'bg-amber-50 border-amber-200' : 'bg-emerald-50 border-emerald-200'}`}>
              <div className="flex items-start gap-3">
                {importResult.errors > 0 ? <AlertCircle className="w-5 h-5 text-amber-600 flex-shrink-0 mt-0.5" /> : <CheckCircle className="w-5 h-5 text-emerald-600 flex-shrink-0 mt-0.5" />}
                <div>
                  <p className={`text-sm font-medium ${importResult.errors > 0 ? 'text-amber-800' : 'text-emerald-800'}`}>
                    Importacao concluida
                  </p>
                  <p className={`text-xs mt-1 ${importResult.errors > 0 ? 'text-amber-600' : 'text-emerald-600'}`}>
                    {importResult.success} registros importados com sucesso
                    {importResult.errors > 0 && ` | ${importResult.errors} registros com erro`}
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* Actions */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <button
              onClick={handleExportExcel}
              disabled={ativos.length === 0}
              className="flex items-center justify-center gap-2 p-4 bg-white border border-slate-200 rounded-xl shadow-sm text-sm font-medium text-slate-700 hover:bg-slate-50 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"
            >
              <Download className="w-4 h-4" />
              Exportar Ativos para Excel
            </button>
            <button
              onClick={handleClearDatabase}
              className="flex items-center justify-center gap-2 p-4 bg-white border border-red-200 rounded-xl shadow-sm text-sm font-medium text-red-600 hover:bg-red-50 transition-colors"
            >
              <Trash2 className="w-4 h-4" />
              Limpar Banco de Dados
            </button>
          </div>
        </main>
      </div>
    );
  }

  // ===================== MAIN PAGE =====================
  return (
    <div className="min-h-screen bg-slate-50">
      <header className="bg-white border-b border-slate-200 sticky top-0 z-30">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <div className="flex items-center gap-3">
              <img src="inv.png" alt="Logo" className="w-10 h-10 object-contain" />
              <div>
                <h1 className="text-lg font-semibold text-slate-900 leading-tight">Controle de Ativos</h1>
                <p className="text-xs text-slate-500">Gestao de patrimonio e equipamentos</p>
              </div>
            </div>
            <div className="flex items-center gap-2">
              <button
                onClick={() => { setPage('admin'); setIsAdminAuth(false); }}
                className="inline-flex items-center gap-2 px-4 py-2.5 bg-slate-900 text-white text-sm font-medium rounded-lg hover:bg-slate-800 transition-all shadow-sm"
              >
                <Shield className="w-4 h-4" />
                Admin
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-6">
        {error && (
          <div className="mb-4 flex items-center gap-2 p-3 bg-red-50 border border-red-200 rounded-lg text-red-700 text-sm">
            <AlertCircle className="w-4 h-4 flex-shrink-0" />{error}
            <button onClick={() => setError(null)} className="ml-auto"><X className="w-4 h-4" /></button>
          </div>
        )}
        {successMessage && (
          <div className="mb-4 flex items-center gap-2 p-3 bg-emerald-50 border border-emerald-200 rounded-lg text-emerald-700 text-sm">
            <CheckCircle className="w-4 h-4 flex-shrink-0" />{successMessage}
          </div>
        )}

        {/* Filter Section */}
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm mb-6">
          <div className="px-5 py-4 border-b border-slate-100">
            <div className="flex items-center justify-between">
              <div className="flex items-center gap-2 text-sm font-medium text-slate-700">
                <Filter className="w-4 h-4 text-emerald-600" />Filtros de Busca
              </div>
                <button onClick={() => { setFiltroPlaca(''); setFiltroLoja(''); setFiltroDescricao(''); setFiltroNumeroBem(''); setFiltroNumeroIncorporacao(''); }} className="text-xs text-slate-400 hover:text-slate-600 transition-colors">Limpar filtros</button>
            </div>
          </div>
          <div className="p-5">
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-4">
              <div>
                <label className="flex items-center gap-1.5 text-xs font-medium text-slate-500 mb-1.5"><Package className="w-3.5 h-3.5" />Numero da Placa</label>
                <div className="relative">
                  <input type="text" value={filtroPlaca} onChange={(e) => setFiltroPlaca(e.target.value)} placeholder="Buscar por placa..." className="w-full pl-3 pr-8 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all" />
                  {filtroPlaca && <button onClick={() => setFiltroPlaca('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><X className="w-4 h-4" /></button>}
                </div>
              </div>
              <div>
                <label className="flex items-center gap-1.5 text-xs font-medium text-slate-500 mb-1.5"><Building2 className="w-3.5 h-3.5" />Numero da Loja</label>
                <div className="relative">
                  <input type="text" value={filtroLoja} onChange={(e) => setFiltroLoja(e.target.value)} placeholder="Buscar por loja..." className="w-full pl-3 pr-8 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all" />
                  {filtroLoja && <button onClick={() => setFiltroLoja('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><X className="w-4 h-4" /></button>}
                </div>
              </div>
              <div>
                  <label className="flex items-center gap-1.5 text-xs font-medium text-slate-500 mb-1.5"><FileText className="w-3.5 h-3.5" />Número do Bem</label>
                  <div className="relative">
                    <input type="text" value={filtroNumeroBem} onChange={(e) => setFiltroNumeroBem(e.target.value)} placeholder="Buscar bem..." className="w-full pl-3 pr-8 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all" />
                    {filtroNumeroBem && <button onClick={() => setFiltroNumeroBem('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><X className="w-4 h-4" /></button>}
                  </div>
                </div>
                <div>
                  <label className="flex items-center gap-1.5 text-xs font-medium text-slate-500 mb-1.5"><FileText className="w-3.5 h-3.5" />Incorporação</label>
                  <div className="relative">
                    <input type="text" value={filtroNumeroIncorporacao} onChange={(e) => setFiltroNumeroIncorporacao(e.target.value)} placeholder="Buscar inc..." className="w-full pl-3 pr-8 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm text-slate-900 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all" />
                    {filtroNumeroIncorporacao && <button onClick={() => setFiltroNumeroIncorporacao('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><X className="w-4 h-4" /></button>}
                  </div>
                </div>
                <div className="lg:col-span-1">
                <label className="flex items-center gap-1.5 text-xs font-medium text-slate-500 mb-1.5"><FileText className="w-3.5 h-3.5" />Descricao do Item</label>
                <div className="relative">
                  <input type="text" value={filtroDescricao} onChange={(e) => setFiltroDescricao(e.target.value)} placeholder="Buscar por descricao..." className="w-full pl-3 pr-8 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm text-slate-900 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 transition-all" />
                  {filtroDescricao && <button onClick={() => setFiltroDescricao('')} className="absolute right-2 top-1/2 -translate-y-1/2 text-slate-400 hover:text-slate-600"><X className="w-4 h-4" /></button>}
                </div>
              </div>
            </div>
          </div>
        </div>

        {/* Table Section */}
        <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
          <div className="px-5 py-4 border-b border-slate-100 flex items-center justify-between">
            <div className="flex items-center gap-2">
              <h2 className="text-sm font-semibold text-slate-800">Ativos Cadastrados</h2>
              <span className="inline-flex items-center px-2 py-0.5 rounded-full text-xs font-medium bg-slate-100 text-slate-600">{ativos.length} {ativos.length === 1 ? 'registro' : 'registros'}</span>
            </div>
            <button onClick={fetchAtivos} disabled={loading} className="inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium text-slate-600 bg-slate-50 border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors">
              <RefreshCw className={`w-3.5 h-3.5 ${loading ? 'animate-spin' : ''}`} />Atualizar
            </button>
          </div>
          {loading ? (
            <div className="flex items-center justify-center py-20"><Loader2 className="w-6 h-6 text-emerald-600 animate-spin" /><span className="ml-2 text-sm text-slate-500">Carregando ativos...</span></div>
          ) : sortedAtivos.length === 0 ? (
            <div className="flex flex-col items-center justify-center py-16 text-slate-400 min-h-[200px]">
              <Package className="w-12 h-12 mb-3 text-slate-300" />
              <p className="text-sm font-medium">Nenhum ativo encontrado</p>
              <p className="text-xs mt-1">Tente ajustar os filtros ou adicione um novo ativo</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs">
                <thead>
                  <tr className="bg-slate-50/50">
                    {visibleColumns.has('placa') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap w-8" onClick={() => handleSort('placa')}><div className="flex items-center gap-1">Placa <SortIcon field="placa" /></div></th>}
                    {visibleColumns.has('numero_loja') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap w-10" onClick={() => handleSort('numero_loja')}><div className="flex items-center gap-1">Loja <SortIcon field="numero_loja" /></div></th>}
                    {visibleColumns.has('descricao') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap" onClick={() => handleSort('descricao')}><div className="flex items-center gap-1">Descrição <SortIcon field="descricao" /></div></th>}
                    {visibleColumns.has('status') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap" onClick={() => handleSort('status')}><div className="flex items-center gap-1">Status <SortIcon field="status" /></div></th>}
                    {visibleColumns.has('categoria') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider whitespace-nowrap">Categoria</th>}
                    {visibleColumns.has('localizacao') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider whitespace-nowrap">Localização</th>}
                    {visibleColumns.has('valor') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap" onClick={() => handleSort('valor')}><div className="flex items-center gap-1">Valor <SortIcon field="valor" /></div></th>}
                    {visibleColumns.has('depr_acum') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap" onClick={() => handleSort('depr_acum')}><div className="flex items-center gap-1">Depr. Acum <SortIcon field="depr_acum" /></div></th>}
                    {visibleColumns.has('saldo_contabil') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap" onClick={() => handleSort('saldo_contabil')}><div className="flex items-center gap-1">Saldo Contábil <SortIcon field="saldo_contabil" /></div></th>}
                    {visibleColumns.has('data_aquisicao') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider cursor-pointer hover:text-slate-700 whitespace-nowrap" onClick={() => handleSort('data_aquisicao')}><div className="flex items-center gap-1">Aquisição <SortIcon field="data_aquisicao" /></div></th>}
                    {visibleColumns.has('observacao') && <th className="px-2 py-3 text-left text-xs font-semibold text-slate-500 uppercase tracking-wider whitespace-nowrap">Observação</th>}
                    <th className="px-2 py-3 text-right text-xs font-semibold text-slate-500 uppercase tracking-wider whitespace-nowrap">Ações</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-100 min-h-[200px]">
                  {displayedAtivos.map((ativo) => (
                    <tr key={ativo.id} className="hover:bg-slate-50/50 transition-colors cursor-pointer" onClick={() => setSelectedAtivo(ativo)}>
                      {visibleColumns.has('placa') && <td className="px-2 py-3 whitespace-nowrap w-8"><span className="font-mono font-medium text-slate-800">{ativo.placa}</span></td>}
                      {visibleColumns.has('numero_loja') && <td className="px-2 py-3 whitespace-nowrap w-10"><span className="inline-flex items-center gap-1 text-slate-700"><Building2 className="w-3.5 h-3.5 text-slate-400" />{ativo.numero_loja}</span></td>}
                      {visibleColumns.has('descricao') && <td className="px-2 py-3 max-w-md"><p className="text-slate-700 truncate" title={ativo.descricao}>{ativo.descricao}</p>{ativo.observacao && <p className="text-[10px] text-slate-400 truncate mt-0.5" title={ativo.observacao}>{ativo.observacao}</p>}</td>}
                      {visibleColumns.has('status') && <td className="px-2 py-3 whitespace-nowrap"><span className={`inline-flex px-2 py-0.5 text-xs font-medium rounded-full border ${statusColors[ativo.status] || 'bg-slate-50 text-slate-600 border-slate-200'}`}>{ativo.status}</span></td>}
                      {visibleColumns.has('categoria') && <td className="px-2 py-3 text-slate-600 whitespace-nowrap">{ativo.categoria || '-'}</td>}
                      {visibleColumns.has('localizacao') && <td className="px-2 py-3 text-slate-600 whitespace-nowrap">{ativo.localizacao || '-'}</td>}
                      {visibleColumns.has('valor') && <td className="px-2 py-3 font-mono text-slate-700 whitespace-nowrap">{formatCurrency(ativo.valor)}</td>}
                      {visibleColumns.has('depr_acum') && <td className="px-2 py-3 font-mono text-slate-700 whitespace-nowrap">{formatCurrency(ativo.depr_acum)}</td>}
                      {visibleColumns.has('saldo_contabil') && <td className="px-2 py-3 font-mono text-slate-700 whitespace-nowrap">{formatCurrency(ativo.saldo_contabil)}</td>}
                      {visibleColumns.has('data_aquisicao') && <td className="px-2 py-3 text-slate-500 whitespace-nowrap">{formatDate(ativo.data_aquisicao)}</td>}
                      {visibleColumns.has('observacao') && <td className="px-2 py-3 text-slate-600 max-w-xs truncate">{ativo.observacao || '-'}</td>}
                      <td className="px-2 py-3 text-right">
                        <div className="inline-flex items-center gap-1">
                          <button onClick={(e) => { e.stopPropagation(); openEditModal(ativo); }} className="p-1.5 text-slate-400 hover:text-emerald-600 hover:bg-emerald-50 rounded-md transition-colors" title="Editar"><Edit3 className="w-4 h-4" /></button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {sortedAtivos.length > displayLimit && (
                <div className="px-5 py-4 border-t border-slate-100 flex justify-center gap-3">
                  <button
                    onClick={() => setDisplayLimit(prev => prev + 10)}
                    className="px-4 py-2 text-sm font-medium text-slate-600 bg-slate-50 border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors"
                  >
                    Carregar mais
                  </button>
                  <button
                    onClick={() => setDisplayLimit(sortedAtivos.length)}
                    className="px-4 py-2 text-sm font-medium text-white bg-slate-600 rounded-lg hover:bg-slate-700 transition-colors shadow-sm"
                  >
                    Carregar tudo
                  </button>
                </div>
              )}
              {sortedAtivos.length > 0 && displayLimit > sortedAtivos.length && (
                <div className="px-5 py-4 border-t border-slate-100 text-center text-sm text-slate-500">
                  Todos os {sortedAtivos.length} ativos carregados.
                </div>
              )}
            </div>
          )}
        </div>
      </main>

      {/* Add Modal */}
      {showAddModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl mx-4 overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between">
              <h3 className="text-base font-semibold text-slate-900">Novo Ativo</h3>
              <button onClick={() => setShowAddModal(false)} className="text-slate-400 hover:text-slate-600"><X className="w-5 h-5" /></button>
            </div>
            <div className="p-6 space-y-4 max-h-[70vh] overflow-y-auto">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Placa *</label><input type="text" value={newAtivo.placa} onChange={(e) => setNewAtivo({ ...newAtivo, placa: e.target.value })} placeholder="ABC-1234" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Loja *</label><input type="text" value={newAtivo.numero_loja} onChange={(e) => setNewAtivo({ ...newAtivo, numero_loja: e.target.value })} placeholder="LOJA-01" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
              </div>
              <div><label className="block text-xs font-medium text-slate-500 mb-1">Descricao *</label><input type="text" value={newAtivo.descricao} onChange={(e) => setNewAtivo({ ...newAtivo, descricao: e.target.value })} placeholder="Descricao do ativo" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Status</label><select value={newAtivo.status} onChange={(e) => setNewAtivo({ ...newAtivo, status: e.target.value })} className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500"><option value="ativo">Ativo</option><option value="inativo">Inativo</option><option value="manutencao">Manutencao</option></select></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Categoria</label><input type="text" value={newAtivo.categoria} onChange={(e) => setNewAtivo({ ...newAtivo, categoria: e.target.value })} placeholder="Informatica" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Localizacao</label><input type="text" value={newAtivo.localizacao} onChange={(e) => setNewAtivo({ ...newAtivo, localizacao: e.target.value })} placeholder="Sala 101" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Data Aquisicao</label><input type="date" value={newAtivo.data_aquisicao} onChange={(e) => setNewAtivo({ ...newAtivo, data_aquisicao: e.target.value })} className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
              </div>
              <div><label className="block text-xs font-medium text-slate-500 mb-1">Valor (R$)</label><input type="number" step="0.01" value={newAtivo.valor} onChange={(e) => setNewAtivo({ ...newAtivo, valor: e.target.value })} placeholder="0.00" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Depr. Acum (R$)</label><input type="number" step="0.01" value={newAtivo.depr_acum} onChange={(e) => setNewAtivo({ ...newAtivo, depr_acum: e.target.value })} placeholder="0.00" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Saldo Contabil (R$)</label><input type="number" step="0.01" value={newAtivo.saldo_contabil} onChange={(e) => setNewAtivo({ ...newAtivo, saldo_contabil: e.target.value })} placeholder="0.00" className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
              </div>
              <div><label className="block text-xs font-medium text-slate-500 mb-1">Observacao</label><textarea value={newAtivo.observacao} onChange={(e) => setNewAtivo({ ...newAtivo, observacao: e.target.value })} rows={2} placeholder="Notas adicionais..." className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 resize-none" /></div>
            </div>
            <div className="px-6 py-4 border-t border-slate-100 flex items-center justify-end gap-3">
              <button onClick={() => setShowAddModal(false)} className="px-4 py-2 text-sm font-medium text-slate-600 bg-slate-50 border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors">Cancelar</button>
              <button onClick={handleAddAtivo} disabled={!newAtivo.placa || !newAtivo.numero_loja || !newAtivo.descricao} className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-white bg-emerald-600 rounded-lg hover:bg-emerald-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed"><Save className="w-4 h-4" />Salvar</button>
            </div>
          </div>
        </div>
      )}

      {/* Edit Modal */}
      {editingAtivo && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl mx-4 overflow-hidden">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between">
              <h3 className="text-base font-semibold text-slate-900">Editar Ativo</h3>
              <button onClick={() => setEditingAtivo(null)} className="text-slate-400 hover:text-slate-600"><X className="w-5 h-5" /></button>
            </div>
            <div className="p-6 space-y-4 max-h-[70vh] overflow-y-auto">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Placa</label><input type="text" value={editForm.placa} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Loja</label><input type="text" value={editForm.numero_loja} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
              </div>
              <div><label className="block text-xs font-medium text-slate-500 mb-1">Descrição</label><input type="text" value={editForm.descricao} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Status</label><select value={editForm.status} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed appearance-none"><option value="ativo">Ativo</option><option value="inativo">Inativo</option><option value="manutencao">Manutenção</option></select></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Categoria</label><input type="text" value={editForm.categoria} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
              </div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Localização</label><input type="text" value={editForm.localizacao} onChange={(e) => setEditForm({ ...editForm, localizacao: e.target.value })} className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500" /></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Data Aquisição</label><input type="date" value={editForm.data_aquisicao} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
              </div>
              <div><label className="block text-xs font-medium text-slate-500 mb-1">Valor (R$)</label><input type="number" value={editForm.valor} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
              <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Depr. Acum (R$)</label><input type="number" value={editForm.depr_acum} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
                <div><label className="block text-xs font-medium text-slate-500 mb-1">Saldo Contábil (R$)</label><input type="number" value={editForm.saldo_contabil} disabled className="w-full px-3 py-2.5 bg-slate-100 border border-slate-200 rounded-lg text-sm cursor-not-allowed" /></div>
              </div>
              <div><label className="block text-xs font-medium text-slate-500 mb-1">Observacao</label><textarea value={editForm.observacao} onChange={(e) => setEditForm({ ...editForm, observacao: e.target.value })} rows={2} className="w-full px-3 py-2.5 bg-slate-50 border border-slate-200 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-emerald-500/20 focus:border-emerald-500 resize-none" /></div>
            </div>
            <div className="px-6 py-4 border-t border-slate-100 flex items-center justify-end gap-3">
              <button onClick={() => setEditingAtivo(null)} className="px-4 py-2 text-sm font-medium text-slate-600 bg-slate-50 border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors">Cancelar</button>
              <button onClick={handleEditAtivo} className="inline-flex items-center gap-1.5 px-4 py-2 text-sm font-medium text-white bg-emerald-600 rounded-lg hover:bg-emerald-700 transition-colors"><Save className="w-4 h-4" />Atualizar</button>
            </div>
          </div>
        </div>
      )}

      {/* Delete Confirmation */}
      {deletingId && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-sm mx-4 overflow-hidden">
            <div className="p-6 text-center">
              <div className="w-12 h-12 mx-auto mb-4 bg-red-50 rounded-full flex items-center justify-center"><Trash2 className="w-6 h-6 text-red-600" /></div>
              <h3 className="text-base font-semibold text-slate-900">Confirmar Exclusao</h3>
              <p className="text-sm text-slate-500 mt-1">Deseja realmente excluir este ativo? Esta acao nao pode ser desfeita.</p>
            </div>
            <div className="px-6 py-4 border-t border-slate-100 flex items-center justify-end gap-3">
              <button onClick={() => setDeletingId(null)} className="px-4 py-2 text-sm font-medium text-slate-600 bg-slate-50 border border-slate-200 rounded-lg hover:bg-slate-100 transition-colors">Cancelar</button>
              <button onClick={() => handleDeleteAtivo(deletingId)} className="px-4 py-2 text-sm font-medium text-white bg-red-600 rounded-lg hover:bg-red-700 transition-colors">Excluir</button>
            </div>
          </div>
        </div>
      )}

      {/* Details Modal */}
      {selectedAtivo && (
        <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/40 backdrop-blur-sm p-4">
          <div className="bg-white rounded-2xl shadow-xl w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh]">
            <div className="px-6 py-4 border-b border-slate-100 flex items-center justify-between bg-slate-50/50">
              <div className="flex items-center gap-2">
                <div className="w-8 h-8 bg-emerald-100 rounded-lg flex items-center justify-center">
                  <Package className="w-4 h-4 text-emerald-600" />
                </div>
                <h3 className="text-base font-semibold text-slate-900">Detalhes do Ativo</h3>
              </div>
              <button onClick={() => setSelectedAtivo(null)} className="p-2 hover:bg-slate-200 rounded-full transition-colors">
                <X className="w-5 h-5 text-slate-400" />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto">
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                <div className="space-y-4">
                  <div>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Identificação</p>
                    <div className="bg-slate-50 p-3 rounded-lg border border-slate-100">
                      <p className="text-xs text-slate-500 mb-0.5">Placa</p>
                      <p className="text-sm font-mono font-bold text-slate-900">{selectedAtivo.placa}</p>
                    </div>
                  </div>
                  
                  <div>
                    <div className="bg-slate-50 p-3 rounded-lg border border-slate-100">
                      <p className="text-xs text-slate-500 mb-0.5">Loja</p>
                      <p className="text-sm font-semibold text-slate-900">{selectedAtivo.numero_loja}</p>
                    </div>
                  </div>

                  <div>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Status e Categoria</p>
                    <div className="flex gap-3">
                      <div className="flex-1 bg-slate-50 p-3 rounded-lg border border-slate-100">
                        <p className="text-xs text-slate-500 mb-1">Status</p>
                        <span className={`inline-flex px-2 py-0.5 text-xs font-medium rounded-full border ${statusColors[selectedAtivo.status] || 'bg-slate-50 text-slate-600 border-slate-200'}`}>
                          {selectedAtivo.status}
                        </span>
                      </div>
                      <div className="flex-1 bg-slate-50 p-3 rounded-lg border border-slate-100">
                        <p className="text-xs text-slate-500 mb-0.5">Categoria</p>
                        <p className="text-sm text-slate-900">{selectedAtivo.categoria || '-'}</p>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="space-y-4">
                  <div>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Valores</p>
                    <div className="bg-slate-50 p-3 rounded-lg border border-slate-100 space-y-3">
                      <div>
                        <p className="text-xs text-slate-500">Valor de Aquisição</p>
                        <p className="text-sm font-mono font-bold text-emerald-600">{formatCurrency(selectedAtivo.valor)}</p>
                      </div>
                      <div className="pt-2 border-t border-slate-200/60">
                        <p className="text-xs text-slate-500">Depreciação Acumulada</p>
                        <p className="text-sm font-mono text-slate-600">{formatCurrency(selectedAtivo.depr_acum)}</p>
                      </div>
                      <div className="pt-2 border-t border-slate-200/60">
                        <p className="text-xs text-slate-500">Saldo Contábil</p>
                        <p className="text-sm font-mono font-bold text-slate-900">{formatCurrency(selectedAtivo.saldo_contabil)}</p>
                      </div>
                    </div>
                  </div>
                  
                  <div>
                    <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Logística</p>
                    <div className="bg-slate-50 p-3 rounded-lg border border-slate-100">
                      <p className="text-xs text-slate-500 mb-0.5">Localização</p>
                      <p className="text-sm text-slate-900">{selectedAtivo.localizacao || '-'}</p>
                      <div className="mt-2 pt-2 border-t border-slate-200/60">
                        <p className="text-xs text-slate-500 mb-0.5">Data de Aquisição</p>
                        <p className="text-sm text-slate-900">{formatDate(selectedAtivo.data_aquisicao)}</p>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

              <div className="mt-6">
                <p className="text-[10px] font-bold text-slate-400 uppercase tracking-wider mb-1">Descrição</p>
                <div className="bg-slate-50 p-4 rounded-lg border border-slate-100">
                  <p className="text-sm text-slate-800 leading-relaxed">{selectedAtivo.descricao}</p>
                </div>
              </div>
            </div>
            
            <div className="px-6 py-4 border-t border-slate-100 bg-slate-50/50 flex justify-end">
              <button onClick={() => setSelectedAtivo(null)} className="px-6 py-2 bg-slate-900 text-white text-sm font-semibold rounded-lg hover:bg-slate-800 transition-colors shadow-sm">Fechar</button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

export default App;
