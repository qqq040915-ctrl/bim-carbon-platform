import { useEffect, useMemo, useRef, useState } from 'react';
import { createPortal } from 'react-dom';
import * as XLSX from 'xlsx';
import { Bar, Pie } from 'react-chartjs-2';
import {
  ArcElement,
  BarElement,
  CategoryScale,
  Chart as ChartJS,
  Legend,
  LinearScale,
  Tooltip,
} from 'chart.js';
import './App.css';

ChartJS.register(CategoryScale, LinearScale, BarElement, ArcElement, Tooltip, Legend);

const CHART_COLORS = ['#315f8c', '#2f855a', '#b7791f', '#805ad5', '#64748b', '#319795', '#dd6b20', '#9f7aea'];
const PROJECT_STORAGE_KEY = 'projectDatabase';
const ACTIVE_PROJECT_STORAGE_KEY = 'activeProjectId';
const MAPPING_STORAGE_KEY = 'materialMappingDatabase';
const HISTORY_STORAGE_KEY = 'bim-a1-a3-analysis-history';
const BIM_IMPORTS_STORAGE_KEY = 'bim-carbon-bim-imports';
const BIM_ROWS_STORAGE_KEY = 'bim-carbon-bim-rows';
const MATERIAL_PLANS_STORAGE_KEY = 'bim-carbon-material-plans';
const CALCULATION_STORAGE_KEY = 'bim-carbon-calculation';
const CLOUD_SYNC_ENDPOINT = '/api/sync';

const NAV_ITEMS = [
  { id: 'project-info', label: '📁 專案資訊' },
  { id: 'bim-import', label: '📂 BIM資料匯入' },
  { id: 'bim-check', label: '🔍 BIM資料檢查' },
  { id: 'mapping-center', label: '🧱 材料對應中心' },
  { id: 'plan-management', label: '📊 材料方案管理' },
  { id: 'carbon-analysis', label: '🌱 A1-A3碳排分析' },
  { id: 'history-results', label: '📈 歷史分析結果' },
  { id: 'report-export', label: '📄 報表匯出' },
];

const VOLUME_COLUMNS = ['體積', 'Volume', 'volume', '體積(m³)', '體積(m3)', '體積 m³', '體積 m3'];
const MATERIAL_COLUMNS = ['結構材料', 'Structural Material', 'Material', 'material', 'Material Name', 'material name', '材料', '材質', '材料名稱', 'BIM材料名稱'];
const CATEGORY_COLUMNS = ['Category', 'category', '類別', '構件名稱', '構件'];
const FAMILY_COLUMNS = ['Family', 'family', '族群', '系列'];
const TYPE_COLUMNS = ['Type', 'type', '類型', '型式'];
const QUANTITY_COLUMNS = ['數量', 'Count', 'count', 'Quantity', 'quantity'];

const FILE_COMPONENT_RULES = [
  { type: '柱', keywords: ['結構柱明細表', '柱明細表'] },
  { type: '梁', keywords: ['結構構架明細表', '構架明細表', '梁明細表'] },
  { type: '樓板', keywords: ['樓板明細表'] },
  { type: '牆', keywords: ['牆明細表', '墙明细表'] },
];

const MATERIAL_RECOMMENDATION_RULES = [
  { category: '混凝土', keywords: ['混凝土', '現場澆注', 'concrete'], targets: ['預拌混凝土', '混凝土'], reason: '原始材料含混凝土或現場澆注，系統搜尋混凝土相關標準材料。' },
  { category: '鋼材', keywords: ['鋼筋', '鋼材', '鋼，', '鋼,', 'steel'], targets: ['熱軋鋼材', '鋼材', '鋼筋', '鋼'], reason: '原始材料含鋼，系統搜尋鋼材、熱軋鋼材或鋼筋相關項目。' },
  { category: '鋁材', keywords: ['鋁擠型', '鋁材', '鋁', 'aluminum', 'aluminium'], targets: ['鋁擠型', '鋁材', '鋁'], reason: '原始材料含鋁，系統搜尋鋁材或鋁擠型相關項目。' },
  { category: '玻璃', keywords: ['low-e', 'lowe', '玻璃', 'glass'], targets: ['Low-E玻璃', '玻璃'], reason: '原始材料含玻璃或 Low-E，系統搜尋玻璃相關項目。' },
  { category: '石膏板', keywords: ['石膏板', 'gypsum'], targets: ['石膏板'], reason: '原始材料含石膏板，系統搜尋同類標準材料。' },
  { category: '纖維水泥板', keywords: ['纖維水泥板', 'fiber cement'], targets: ['纖維水泥板'], reason: '原始材料為纖維水泥板，系統優先尋找同名標準材料。' },
  { category: '金屬構件', keywords: ['金屬', 'metal'], targets: ['金屬', '鋼材', '鋁材'], reason: '原始材料含金屬，系統搜尋金屬構件相關標準材料。' },
  { category: '門窗材料', keywords: ['門', '窗', 'door', 'window'], targets: ['門窗', '鋁材', '玻璃'], reason: '原始材料或構件名稱含門窗，系統搜尋門窗、鋁材或玻璃相關項目。' },
  { category: '裝修材料', keywords: ['裝修', '飾面', '天花', '地坪', 'finish'], targets: ['裝修', '飾面', '地坪', '天花'], reason: '原始材料含裝修語彙，系統搜尋裝修材料相關項目。' },
];

function today() {
  return new Date().toISOString().slice(0, 10);
}

function emptyProjectInfo() {
  return {
    projectName: '',
    engineeringName: '',
    location: '',
    buildingUse: '',
    structureType: 'RC',
    floors: '',
    buildingArea: '',
    createdDate: today(),
    note: '',
  };
}

function normalizeProject(project) {
  return {
    project_id: project.project_id ?? project.projectId ?? '',
    projectName: project.projectName ?? project.project_name ?? '',
    engineeringName: project.engineeringName ?? project.engineering_name ?? '',
    location: project.location ?? '',
    buildingUse: project.buildingUse ?? project.building_use ?? '',
    structureType: project.structureType ?? project.structure_type ?? 'RC',
    floors: project.floors ?? '',
    buildingArea: project.buildingArea ?? project.building_area ?? '',
    createdDate: project.createdDate ?? project.created_date ?? today(),
    note: project.note ?? '',
  };
}

function generateProjectId(projects) {
  const datePart = today().replaceAll('-', '');
  const prefix = `P${datePart}`;
  const sequence =
    projects.filter((project) => String(project.project_id ?? '').startsWith(prefix)).length + 1;
  return `${prefix}${String(sequence).padStart(3, '0')}`;
}

function toNumber(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function parseNumber(value) {
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const match = String(value ?? '').replace(/,/g, '').match(/-?\d+(\.\d+)?/);
  return match ? Number(match[0]) : 0;
}

function formatNumber(value, digits = 2) {
  return toNumber(value).toLocaleString(undefined, {
    minimumFractionDigits: digits,
    maximumFractionDigits: digits,
  });
}

function safeFileName(value) {
  return String(value || '未命名專案')
    .trim()
    .replace(/[\\/:*?"<>|]/g, '_')
    .replace(/\s+/g, '_');
}

function normalizeName(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/\s+/g, '')
    .replace(/[_\-()（）]/g, '');
}

function normalizeMaterialText(value) {
  return String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/^\d+[-－]/, '')
    .replace(/[，,、:：;；/\\()（）[\]{}]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function materialTokens(value) {
  const normalized = normalizeMaterialText(value);
  const asciiTokens = normalized.match(/[a-z0-9]+/g) ?? [];
  const chineseTokens = MATERIAL_RECOMMENDATION_RULES.flatMap((rule) => [...rule.keywords, ...rule.targets])
    .filter((keyword) => normalizeName(normalized).includes(normalizeName(keyword)));
  return [...new Set([...asciiTokens, ...chineseTokens].filter((token) => normalizeName(token).length >= 1))];
}

function findColumn(headers, candidates) {
  const targets = candidates.map(normalizeName);
  return headers.find((header) => targets.includes(normalizeName(header))) ?? '';
}

function readWorkbook(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const workbook = XLSX.read(event.target.result, { type: 'array' });
        resolve(workbook.Sheets[workbook.SheetNames[0]]);
      } catch (error) {
        reject(error);
      }
    };
    reader.onerror = () => reject(new Error('檔案讀取失敗'));
    reader.readAsArrayBuffer(file);
  });
}

function worksheetToObjects(worksheet) {
  const matrix = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
  const headerIndex = matrix.findIndex((row) => {
    const headers = row.map((cell) => String(cell).trim());
    return findColumn(headers, VOLUME_COLUMNS) || findColumn(headers, MATERIAL_COLUMNS) || findColumn(headers, CATEGORY_COLUMNS);
  });

  if (headerIndex === -1) return { rows: [], columns: {} };

  const headers = matrix[headerIndex].map((cell, index) => String(cell).trim() || `欄位${index + 1}`);
  const columns = {
    category: findColumn(headers, CATEGORY_COLUMNS),
    family: findColumn(headers, FAMILY_COLUMNS),
    type: findColumn(headers, TYPE_COLUMNS),
    material: findColumn(headers, MATERIAL_COLUMNS),
    volume: findColumn(headers, VOLUME_COLUMNS),
    quantity: findColumn(headers, QUANTITY_COLUMNS),
  };
  const rows = matrix
    .slice(headerIndex + 1)
    .filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''))
    .map((row) => Object.fromEntries(headers.map((header, index) => [header, row[index] ?? ''])));

  return { rows, columns };
}

function coefficientRowsFromWorksheet(worksheet) {
  return XLSX.utils
    .sheet_to_json(worksheet, { defval: '' })
    .map((row, index) => ({
      id: `${row.item_name || row.name || 'item'}-${index}`,
      itemName: String(row.item_name ?? row.name ?? '').trim(),
      coe: parseNumber(row.coe),
      unit: String(row.unit ?? '').trim() || 'kgCO2e/m3',
      source: String(row.departmentname ?? row.source ?? '').trim(),
      year: String(row.announcementyear ?? '').trim(),
    }))
    .filter((row) => row.itemName && row.coe > 0);
}

function getComponentTypeFromFileName(fileName) {
  const normalizedFileName = normalizeName(fileName);
  const rule = FILE_COMPONENT_RULES.find((item) => item.keywords.some((keyword) => normalizedFileName.includes(normalizeName(keyword))));
  return rule?.type ?? '其他構件';
}

function classifyStructuralMaterial(materialName) {
  const normalizedMaterial = normalizeName(materialName);
  const rule = MATERIAL_RECOMMENDATION_RULES.find((item) =>
    item.keywords.some((keyword) => normalizedMaterial.includes(normalizeName(keyword)))
  );
  return rule?.category ?? '待確認材料';
}

function matchesMaterialCategory(itemName, category) {
  const materialCategory = classifyStructuralMaterial(itemName);
  if (category === '全部') return true;
  if (category === '板材') return ['石膏板', '纖維水泥板'].includes(materialCategory);
  if (category === '其他') return materialCategory === '待確認材料';
  return materialCategory === category;
}

function filterMaterialOptions(databaseRows, query, category, limit = 20) {
  const normalizedQuery = normalizeName(query);
  const results = databaseRows
    .filter((row) => matchesMaterialCategory(row.itemName, category))
    .map((row) => ({
      ...row,
      searchScore: normalizedQuery ? scoreMaterialMatch(query, row) : 1,
      directMatch: normalizedQuery ? normalizeName(row.itemName).includes(normalizedQuery) : true,
    }))
    .filter((row) => !normalizedQuery || row.directMatch || row.searchScore >= 35)
    .sort((a, b) => {
      if (b.searchScore !== a.searchScore) return b.searchScore - a.searchScore;
      return a.itemName.localeCompare(b.itemName, 'zh-Hant');
    });
  return Number.isFinite(limit) ? results.slice(0, limit) : results;
}

function scoreMaterialMatch(keyword, row) {
  const key = normalizeName(normalizeMaterialText(keyword));
  const item = normalizeName(row.itemName);
  if (!key || !item) return 0;
  if (key === item) return 100;

  let score = 0;
  if (item.includes(key) || key.includes(item)) score += 60;

  const tokens = materialTokens(keyword);
  const matchedTokens = tokens.filter((token) => item.includes(normalizeName(token)));
  score += Math.min(matchedTokens.length * 22, 55);

  const keyChars = [...new Set([...key])];
  const overlap = keyChars.filter((char) => item.includes(char)).length;
  if (keyChars.length > 0) score += Math.min((overlap / keyChars.length) * 30, 30);

  MATERIAL_RECOMMENDATION_RULES.forEach((rule) => {
    const hasSourceKeyword = rule.keywords.some((keywordItem) => key.includes(normalizeName(keywordItem)));
    const hasTargetKeyword = rule.targets.some((target) => item.includes(normalizeName(target)));
    if (hasSourceKeyword && hasTargetKeyword) score += 42;
  });

  const strength = key.match(/(210|240|280|300|350|400|420|500)/)?.[1];
  if (strength && item.includes(strength)) score += 45;
  return Math.min(Math.round(score), 99);
}

function findRuleMaterial(structuralMaterial, databaseRows) {
  const normalizedMaterial = normalizeName(structuralMaterial);
  const rule = MATERIAL_RECOMMENDATION_RULES.find((item) =>
    item.keywords.some((keyword) => normalizedMaterial.includes(normalizeName(keyword)))
  );
  if (!rule) return null;
  const candidates = databaseRows
    .filter((row) => rule.targets.some((target) => normalizeName(row.itemName).includes(normalizeName(target))))
    .map((row) => ({ ...row, similarity: Math.max(scoreMaterialMatch(structuralMaterial, row), 72) }))
    .sort((a, b) => b.similarity - a.similarity);
  const preferred =
    candidates.find((row) => normalizeName(structuralMaterial).includes('混凝土') && normalizeName(row.itemName).includes('350')) ??
    candidates[0];
  return preferred ? { ...preferred, similarity: 88, reason: rule.reason } : null;
}

function recommendMaterial(searchText, databaseRows) {
  const exact = databaseRows.find((row) => normalizeName(row.itemName) === normalizeName(normalizeMaterialText(searchText)));
  if (exact) return { ...exact, similarity: 100, reason: '完全比對 BIM 原始材料名稱與 item_name' };

  const rows = databaseRows
    .map((row) => ({ ...row, similarity: scoreMaterialMatch(searchText, row) }))
    .filter((row) => row.similarity >= 45)
    .sort((a, b) => b.similarity - a.similarity);

  if (rows[0]) return { ...rows[0], reason: '依 BIM 原始材料名稱進行關鍵字與模糊比對' };

  const ruleMaterial = findRuleMaterial(searchText, databaseRows);
  if (ruleMaterial) return ruleMaterial;
  return null;
}

function normalizeMapping(mapping) {
  return {
    id: mapping.id || crypto.randomUUID(),
    componentType: mapping.componentType ?? mapping.component_type ?? '',
    bimMaterialName: mapping.bimMaterialName ?? mapping.bim_material_name ?? '',
    itemName: mapping.itemName ?? mapping.item_name ?? '',
    carbonFactor: toNumber(mapping.carbonFactor ?? mapping.carbon_factor ?? mapping.coefficient),
    unit: mapping.unit || 'kgCO2e/m3',
    source: mapping.source || '',
    status: mapping.status || 'confirmed',
    note: mapping.note || '',
    project_id: mapping.project_id ?? mapping.projectId ?? '',
    createdAt: mapping.createdAt ?? mapping.created_at ?? new Date().toISOString(),
  };
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

function SearchableMaterialSelector({ category, databaseRows, disabled, onSelect, value }) {
  const inputRef = useRef(null);
  const [inputValue, setInputValue] = useState(value || '');
  const [isOpen, setIsOpen] = useState(false);
  const [activeIndex, setActiveIndex] = useState(0);
  const [popupStyle, setPopupStyle] = useState({ left: 0, top: 0, width: 420 });

  const allOptions = useMemo(
    () => filterMaterialOptions(databaseRows, inputValue, category, Infinity),
    [category, databaseRows, inputValue]
  );
  const options = useMemo(() => allOptions.slice(0, 20), [allOptions]);

  function updatePopupPosition() {
    const input = inputRef.current;
    if (!input) return;
    const rect = input.getBoundingClientRect();
    const width = Math.min(Math.max(rect.width, 420), window.innerWidth - 24);
    const left = Math.max(12, Math.min(rect.left, window.innerWidth - width - 12));
    const top = Math.min(rect.bottom + 6, window.innerHeight - 90);
    setPopupStyle({ left, top, width });
  }

  function openSuggestions() {
    updatePopupPosition();
    setIsOpen(true);
  }

  useEffect(() => {
    if (!isOpen) return undefined;
    const handleViewportChange = () => updatePopupPosition();
    window.addEventListener('resize', handleViewportChange);
    window.addEventListener('scroll', handleViewportChange, true);
    return () => {
      window.removeEventListener('resize', handleViewportChange);
      window.removeEventListener('scroll', handleViewportChange, true);
    };
  }, [isOpen]);

  function selectMaterial(row) {
    onSelect(row);
    setInputValue(row.itemName);
    setIsOpen(false);
    setActiveIndex(0);
  }

  function handleKeyDown(event) {
    if (!isOpen && ['ArrowDown', 'ArrowUp'].includes(event.key)) {
      openSuggestions();
      return;
    }
    if (event.key === 'ArrowDown') {
      event.preventDefault();
      setActiveIndex((current) => Math.min(current + 1, Math.max(options.length - 1, 0)));
    }
    if (event.key === 'ArrowUp') {
      event.preventDefault();
      setActiveIndex((current) => Math.max(current - 1, 0));
    }
    if (event.key === 'Enter' && isOpen && options[activeIndex]) {
      event.preventDefault();
      selectMaterial(options[activeIndex]);
    }
    if (event.key === 'Escape') {
      setIsOpen(false);
    }
  }

  return (
    <div className="material-selector">
      <input
        disabled={disabled}
        ref={inputRef}
        onBlur={() => window.setTimeout(() => setIsOpen(false), 120)}
        onChange={(event) => {
          setInputValue(event.target.value);
          setActiveIndex(0);
          openSuggestions();
        }}
        onFocus={openSuggestions}
        onKeyDown={handleKeyDown}
        placeholder="搜尋材料名稱"
        value={inputValue}
      />
      {isOpen && !disabled && createPortal(
        <div className="material-suggestions" role="listbox" style={popupStyle}>
          <div className="material-suggestions-meta">找到 {allOptions.length} 筆材料</div>
          {options.map((row, index) => (
            <button
              className={index === activeIndex ? 'active' : ''}
              key={row.id}
              onMouseDown={(event) => {
                event.preventDefault();
                selectMaterial(row);
              }}
              type="button"
            >
              <span>{row.itemName}</span>
              <small>{formatNumber(row.coe, 4)} {row.unit}</small>
              <small>資料來源：{row.source || '未提供'}</small>
            </button>
          ))}
          {allOptions.length > 20 && <p>還有更多結果，請繼續輸入關鍵字。</p>}
          {options.length === 0 && <p>找不到符合材料，請調整關鍵字或分類。</p>}
        </div>,
        document.body
      )}
    </div>
  );
}

function App() {
  const [activeSection, setActiveSection] = useState('project-info');
  const [projects, setProjects] = useState([]);
  const [activeProject, setActiveProject] = useState(null);
  const [isProjectFormOpen, setIsProjectFormOpen] = useState(false);
  const [editingProjectId, setEditingProjectId] = useState('');
  const [projectInfo, setProjectInfo] = useState(emptyProjectInfo);
  const [databaseRows, setDatabaseRows] = useState([]);
  const [databaseMessage, setDatabaseMessage] = useState('正在讀取 Preview_Data.csv...');
  const [databaseQuery, setDatabaseQuery] = useState('');
  const [planMaterialCategory, setPlanMaterialCategory] = useState('全部');
  const [imports, setImports] = useState([]);
  const [bimRows, setBimRows] = useState([]);
  const [materialMappings, setMaterialMappings] = useState([]);
  const [materialPlans, setMaterialPlans] = useState([]);
  const [mappingDraft, setMappingDraft] = useState({
    componentType: '',
    bimMaterialName: '',
    itemName: '',
    carbonFactor: '',
    unit: 'kgCO2e/m3',
    source: '',
  });
  const [calculation, setCalculation] = useState(null);
  const [expandedPlanIds, setExpandedPlanIds] = useState([]);
  const [expandedMappingFiles, setExpandedMappingFiles] = useState([]);
  const [history, setHistory] = useState([]);
  const [viewedFileId, setViewedFileId] = useState('');
  const [statusMessage, setStatusMessage] = useState('請先建立並儲存專案資訊，才能進行 BIM 資料匯入與碳排分析。');
  const cloudSyncReadyRef = useRef(false);
  const cloudSyncTimerRef = useRef(null);
  const cloudFallbackRef = useRef(false);

  function loadLocalState() {
    let loadedActiveProject = null;
    try {
      const savedProjects = JSON.parse(localStorage.getItem(PROJECT_STORAGE_KEY) || '[]');
      const normalizedProjects = Array.isArray(savedProjects) ? savedProjects.map(normalizeProject) : [];
      const savedActiveProjectId = localStorage.getItem(ACTIVE_PROJECT_STORAGE_KEY) || '';
      loadedActiveProject = normalizedProjects.find((project) => project.project_id === savedActiveProjectId) ?? null;
      setProjects(normalizedProjects);
      setActiveProject(loadedActiveProject);
      if (loadedActiveProject) {
        setProjectInfo(loadedActiveProject);
        setStatusMessage('目前專案已載入，可以進行 BIM 資料匯入與碳排分析。');
      }
    } catch {
      setProjects([]);
      setActiveProject(null);
    }

    try {
      const savedMappings = JSON.parse(localStorage.getItem(MAPPING_STORAGE_KEY) || '[]');
      const normalizedMappings = Array.isArray(savedMappings) ? savedMappings.map(normalizeMapping) : [];
      setMaterialMappings(
        loadedActiveProject
          ? normalizedMappings.filter((mapping) => !mapping.project_id || mapping.project_id === loadedActiveProject.project_id)
          : []
      );
    } catch {
      setMaterialMappings([]);
    }

    try {
      const savedHistory = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || '[]');
      setHistory(
        Array.isArray(savedHistory) && loadedActiveProject
          ? savedHistory.filter((record) => record.project_id === loadedActiveProject.project_id)
          : []
      );
    } catch {
      setHistory([]);
    }

    try {
      const savedImports = JSON.parse(localStorage.getItem(BIM_IMPORTS_STORAGE_KEY) || '[]');
      const savedBimRows = JSON.parse(localStorage.getItem(BIM_ROWS_STORAGE_KEY) || '[]');
      const savedPlans = JSON.parse(localStorage.getItem(MATERIAL_PLANS_STORAGE_KEY) || '[]');
      const savedCalculation = JSON.parse(localStorage.getItem(CALCULATION_STORAGE_KEY) || 'null');
      setImports(
        Array.isArray(savedImports) && loadedActiveProject
          ? savedImports.filter((item) => !item.project_id || item.project_id === loadedActiveProject.project_id)
          : []
      );
      setBimRows(
        Array.isArray(savedBimRows) && loadedActiveProject
          ? savedBimRows.filter((row) => !row.project_id || row.project_id === loadedActiveProject.project_id)
          : []
      );
      setMaterialPlans(
        Array.isArray(savedPlans) && loadedActiveProject
          ? savedPlans.filter((plan) => !plan.project_id || plan.project_id === loadedActiveProject.project_id)
          : []
      );
      setCalculation(savedCalculation);
    } catch {
      setImports([]);
      setBimRows([]);
      setMaterialPlans([]);
      setCalculation(null);
    }
    return loadedActiveProject;
  }

  async function apiRequest(path, options = {}) {
    const response = await fetch(path, {
      ...options,
      headers: {
        'content-type': 'application/json',
        ...(options.headers ?? {}),
      },
    });
    const data = await response.json();
    if (!response.ok || data.ok === false) {
      throw new Error(data.error || `API request failed: ${response.status}`);
    }
    return data;
  }

  function persistLocalSnapshot(nextState) {
    const projectId = nextState.activeProject?.project_id ?? '';
    const mergeProjectItems = (storageKey, items) => {
      if (nextState.fullSnapshot || !projectId) return items;
      const saved = readStorageJson(storageKey, []);
      return [
        ...saved.filter((item) => item.project_id && item.project_id !== projectId),
        ...items,
      ];
    };
    const nextMappings = mergeProjectItems(MAPPING_STORAGE_KEY, nextState.materialMappings);
    const nextHistory = mergeProjectItems(HISTORY_STORAGE_KEY, nextState.history);
    const nextImports = mergeProjectItems(BIM_IMPORTS_STORAGE_KEY, nextState.imports ?? imports);
    const nextBimRows = mergeProjectItems(BIM_ROWS_STORAGE_KEY, nextState.bimRows ?? bimRows);
    const nextPlans = mergeProjectItems(MATERIAL_PLANS_STORAGE_KEY, nextState.materialPlans ?? materialPlans);
    localStorage.setItem(PROJECT_STORAGE_KEY, JSON.stringify(nextState.projects));
    localStorage.setItem(ACTIVE_PROJECT_STORAGE_KEY, nextState.activeProject?.project_id ?? '');
    localStorage.setItem(MAPPING_STORAGE_KEY, JSON.stringify(nextMappings));
    localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(nextHistory));
    localStorage.setItem(BIM_IMPORTS_STORAGE_KEY, JSON.stringify(nextImports));
    localStorage.setItem(BIM_ROWS_STORAGE_KEY, JSON.stringify(nextBimRows));
    localStorage.setItem(MATERIAL_PLANS_STORAGE_KEY, JSON.stringify(nextPlans));
    localStorage.setItem(CALCULATION_STORAGE_KEY, JSON.stringify(nextState.calculation ?? calculation));
  }

  function applyCloudState(cloudState) {
    const normalizedProjects = Array.isArray(cloudState.projectDatabase)
      ? cloudState.projectDatabase.map(normalizeProject).filter((project) => project.project_id)
      : [];
    const activeProjectId = cloudState.activeProjectId || cloudState.activeProject?.project_id || '';
    const loadedActiveProject =
      normalizedProjects.find((project) => project.project_id === activeProjectId) ??
      normalizedProjects[0] ??
      null;
    const allMappings = Array.isArray(cloudState.materialMappingDatabase)
      ? cloudState.materialMappingDatabase.map(normalizeMapping)
      : [];
    const allHistory = Array.isArray(cloudState.history) ? cloudState.history : [];
    const nextImports = Array.isArray(cloudState.bimImportSummary) ? cloudState.bimImportSummary : [];
    const nextBimRows = Array.isArray(cloudState.bimRows) ? cloudState.bimRows : [];
    const nextPlans = Array.isArray(cloudState.materialPlans) ? cloudState.materialPlans : [];

    setProjects(normalizedProjects);
    setActiveProject(loadedActiveProject);
    setProjectInfo(loadedActiveProject ?? emptyProjectInfo());
    setImports(loadedActiveProject ? nextImports.filter((item) => !item.project_id || item.project_id === loadedActiveProject.project_id) : nextImports);
    setBimRows(loadedActiveProject ? nextBimRows.filter((row) => !row.project_id || row.project_id === loadedActiveProject.project_id) : nextBimRows);
    setMaterialMappings(
      loadedActiveProject
        ? allMappings.filter((mapping) => !mapping.project_id || mapping.project_id === loadedActiveProject.project_id)
        : []
    );
    setMaterialPlans(
      loadedActiveProject ? nextPlans.filter((plan) => !plan.project_id || plan.project_id === loadedActiveProject.project_id) : nextPlans
    );
    setHistory(loadedActiveProject ? filterProjectRecords(allHistory, loadedActiveProject.project_id) : []);
    setCalculation(cloudState.calculation ?? null);
    persistLocalSnapshot({
      projects: normalizedProjects,
      activeProject: loadedActiveProject,
      materialMappings: allMappings,
      history: allHistory,
      imports: nextImports,
      bimRows: nextBimRows,
      materialPlans: nextPlans,
      calculation: cloudState.calculation ?? null,
      fullSnapshot: true,
    });
    if (loadedActiveProject) {
      setStatusMessage('雲端資料已載入，可以進行 BIM 資料匯入與碳排分析。');
    } else {
      setStatusMessage('雲端資料庫已連線，請建立專案後開始分析。');
    }
  }

  function buildCloudPayload() {
    const storedProjects = readStorageJson(PROJECT_STORAGE_KEY, projects);
    const projectId = activeProject?.project_id ?? '';
    const mergeProjectItems = (storageKey, items) => {
      const saved = readStorageJson(storageKey, []);
      if (!projectId) return saved.length ? saved : items;
      return [
        ...saved.filter((item) => item.project_id && item.project_id !== projectId),
        ...items,
      ];
    };
    const mergedMappings = mergeProjectItems(MAPPING_STORAGE_KEY, materialMappings);
    const mergedHistory = mergeProjectItems(HISTORY_STORAGE_KEY, history);
    const mergedImports = mergeProjectItems(BIM_IMPORTS_STORAGE_KEY, imports);
    const mergedBimRows = mergeProjectItems(BIM_ROWS_STORAGE_KEY, bimRows);
    const mergedPlans = mergeProjectItems(MATERIAL_PLANS_STORAGE_KEY, materialPlans);
    return {
      projectDatabase: storedProjects.length ? storedProjects : projects,
      activeProject,
      activeProjectId: activeProject?.project_id ?? '',
      bimImportSummary: mergedImports,
      bimRows: mergedBimRows,
      materialMappingDatabase: mergedMappings,
      materialPlans: mergedPlans,
      history: mergedHistory,
      calculation,
    };
  }

  async function syncCloudData() {
    if (!cloudSyncReadyRef.current || cloudFallbackRef.current) return;
    try {
      await apiRequest(CLOUD_SYNC_ENDPOINT, {
        method: 'POST',
        body: JSON.stringify(buildCloudPayload()),
      });
    } catch {
      cloudFallbackRef.current = true;
      setStatusMessage('D1 雲端資料庫暫時無法連線，已改用瀏覽器本機資料作為備援。');
    }
  }

  useEffect(() => {
    let canceled = false;
    loadLocalState();

    apiRequest(CLOUD_SYNC_ENDPOINT)
      .then((cloudState) => {
        if (canceled) return;
        applyCloudState(cloudState);
        cloudFallbackRef.current = false;
      })
      .catch(() => {
        if (canceled) return;
        cloudFallbackRef.current = true;
        setStatusMessage('D1 雲端資料庫暫時無法連線，已使用瀏覽器本機資料作為備援。');
      })
      .finally(() => {
        if (!canceled) cloudSyncReadyRef.current = true;
      });

    return () => {
      canceled = true;
      if (cloudSyncTimerRef.current) clearTimeout(cloudSyncTimerRef.current);
    };
    // Cloud bootstrap intentionally runs once; subsequent changes are handled by the sync effect below.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    if (!cloudSyncReadyRef.current || cloudFallbackRef.current) return;
    if (cloudSyncTimerRef.current) clearTimeout(cloudSyncTimerRef.current);
    cloudSyncTimerRef.current = setTimeout(() => {
      syncCloudData();
    }, 600);
    // Sync uses a state snapshot from this render and refs to avoid recursive cloud writes.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projects, activeProject, imports, bimRows, materialMappings, materialPlans, history, calculation]);

  useEffect(() => {
    if (!cloudSyncReadyRef.current) return;
    persistLocalSnapshot({
      projects,
      activeProject,
      materialMappings,
      history,
      imports,
      bimRows,
      materialPlans,
      calculation,
    });
    // Local fallback persistence mirrors the same state snapshot without becoming a dependency source.
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [projects, activeProject, imports, bimRows, materialMappings, materialPlans, history, calculation]);

  useEffect(() => {
    fetch('/Preview_Data.csv')
      .then((response) => {
        if (!response.ok) throw new Error('Preview_Data.csv not found');
        return response.arrayBuffer();
      })
      .then((buffer) => {
        const workbook = XLSX.read(buffer, { type: 'array' });
        const rows = coefficientRowsFromWorksheet(workbook.Sheets[workbook.SheetNames[0]]);
        setDatabaseRows(rows);
        setDatabaseMessage(`已載入 ${rows.length} 筆碳排係數資料。`);
      })
      .catch(() => setDatabaseMessage('無法載入 Preview_Data.csv，請確認檔案位於 public 資料夾。'));
  }, []);

  useEffect(() => {
    const observer = new IntersectionObserver(
      (entries) => {
        const visible = entries
          .filter((entry) => entry.isIntersecting)
          .sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
        if (visible) setActiveSection(visible.target.id);
      },
      { rootMargin: '-20% 0px -65% 0px', threshold: [0.15, 0.35] }
    );

    NAV_ITEMS.forEach((item) => {
      const element = document.getElementById(item.id);
      if (element) observer.observe(element);
    });

    return () => observer.disconnect();
  }, [calculation, imports.length, materialPlans.length]);

  const mappingByKey = useMemo(() => {
    const map = new Map();
    materialMappings.forEach((mapping) => {
      if (mapping.bimMaterialName) map.set(normalizeName(mapping.bimMaterialName), mapping);
    });
    return map;
  }, [materialMappings]);

  const mappingCenterRows = useMemo(() => {
    const grouped = new Map();
    bimRows.forEach((row) => {
      const materialName = row.materialName || `${row.componentType}（結構材料待確認）`;
      const key = `${row.fileId}-${row.componentType}-${materialName}`;
      const current = grouped.get(key) ?? {
        id: key,
        fileId: row.fileId,
        fileName: row.fileName,
        componentType: row.componentType,
        bimMaterialName: materialName,
        materialCategory: row.materialCategory,
        count: 0,
        volume: 0,
      };
      current.count += 1;
      current.volume += row.volume;
      grouped.set(key, current);
    });

    return [...grouped.values()].map((item) => {
      const mapping = mappingByKey.get(normalizeName(item.bimMaterialName));
      const recommendation = mapping ? null : recommendMaterial(item.bimMaterialName, databaseRows);
      return {
        ...item,
        mapping,
        recommendation,
        status: mapping ? '已完成對應' : recommendation ? '待確認' : '待確認材料',
      };
    });
  }, [bimRows, databaseRows, mappingByKey]);

  const mappingGroups = useMemo(() => {
    const grouped = new Map();
    mappingCenterRows.forEach((row) => {
      const group = grouped.get(row.fileId) ?? {
        fileId: row.fileId,
        fileName: row.fileName,
        rows: [],
      };
      group.rows.push(row);
      grouped.set(row.fileId, group);
    });

    return [...grouped.values()].map((group) => {
      const completed = group.rows.filter((row) => row.status === '已完成對應').length;
      const pending = group.rows.length - completed;
      return {
        ...group,
        completed,
        materialCount: group.rows.length,
        pending,
        totalVolume: group.rows.reduce((sum, row) => sum + row.volume, 0),
      };
    });
  }, [mappingCenterRows]);

  useEffect(() => {
    setExpandedMappingFiles((current) => {
      if (current.length > 0) return current;
      return mappingGroups.map((group) => group.fileId);
    });
  }, [mappingGroups]);

  const quality = useMemo(() => {
    const total = bimRows.length;
    const missing = {
      type: bimRows.filter((row) => !row.hasType).length,
      material: bimRows.filter((row) => !row.hasMaterial).length,
      volume: bimRows.filter((row) => !row.hasVolume).length,
      quantity: bimRows.filter((row) => !row.hasQuantity).length,
    };
    const missingTotal = Object.values(missing).reduce((sum, count) => sum + count, 0);
    const fieldTotal = total * 4;
    const completionRate = fieldTotal > 0 ? ((fieldTotal - missingTotal) / fieldTotal) * 100 : 0;
    const uniqueMaterialCount = new Set(bimRows.map((row) => row.materialName).filter(Boolean)).size;
    return { total, missing, missingTotal, completionRate, uniqueMaterialCount };
  }, [bimRows]);

  const mappingStats = useMemo(() => {
    const completed = mappingCenterRows.filter((row) => row.status === '已完成對應').length;
    const pending = mappingCenterRows.filter((row) => row.status === '待確認').length;
    const unmapped = mappingCenterRows.filter((row) => row.status === '待確認材料').length;
    const total = mappingCenterRows.length;
    const recommended = mappingCenterRows.filter((row) => row.mapping || row.recommendation).length;
    const recommendationRate = total > 0 ? (recommended / total) * 100 : 0;
    const confirmationRate = total > 0 ? (completed / total) * 100 : 0;
    return { completed, pending, unmapped, total, recommended, recommendationRate, confirmationRate };
  }, [mappingCenterRows]);

  const totalVolume = useMemo(() => bimRows.reduce((sum, row) => sum + row.volume, 0), [bimRows]);
  const filteredDatabaseRows = useMemo(() => {
    const query = databaseQuery.trim().toLowerCase();
    if (!query) return databaseRows.slice(0, 160);
    return databaseRows.filter((row) => row.itemName.toLowerCase().includes(query)).slice(0, 160);
  }, [databaseRows, databaseQuery]);

  const mappingSelectRows = useMemo(() => {
    if (!mappingDraft.itemName || filteredDatabaseRows.some((row) => row.itemName === mappingDraft.itemName)) {
      return filteredDatabaseRows;
    }
    const selected = databaseRows.find((row) => row.itemName === mappingDraft.itemName);
    return selected ? [selected, ...filteredDatabaseRows] : filteredDatabaseRows;
  }, [databaseRows, filteredDatabaseRows, mappingDraft.itemName]);

  const planBarData = useMemo(
    () => ({
      labels: calculation?.planResults.map((plan) => plan.name) ?? [],
      datasets: [
        {
          label: '總碳排 kgCO2e',
          data: calculation?.planResults.map((plan) => plan.totalCarbon) ?? [],
          backgroundColor: '#315f8c',
          borderRadius: 4,
        },
      ],
    }),
    [calculation]
  );

  const chartOptions = {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { position: 'bottom' },
      tooltip: { callbacks: { label: (context) => `${context.label}: ${formatNumber(Number(context.raw))} kgCO2e` } },
    },
  };

  function scrollToSection(sectionId) {
    document.getElementById(sectionId)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    setActiveSection(sectionId);
  }

  function updateProject(field, value) {
    setProjectInfo((current) => ({ ...current, [field]: value }));
  }

  function readStorageJson(key, fallback = []) {
    try {
      const parsed = JSON.parse(localStorage.getItem(key) || '');
      return Array.isArray(parsed) ? parsed : fallback;
    } catch {
      return fallback;
    }
  }

  function exportProjectBackup() {
    const projectName = activeProject?.projectName || projectInfo.projectName || '未命名專案';
    const storedProjects = readStorageJson(PROJECT_STORAGE_KEY, projects);
    const storedMappings = readStorageJson(MAPPING_STORAGE_KEY, materialMappings);
    const storedHistory = readStorageJson(HISTORY_STORAGE_KEY, history);
    const backup = {
      backupType: 'bim-carbon-platform-project-backup',
      version: 1,
      exportedAt: new Date().toISOString(),
      projectDatabase: storedProjects.length ? storedProjects : projects,
      activeProject,
      activeProjectId: activeProject?.project_id ?? '',
      bimImportSummary: imports,
      bimRows,
      materialMappingDatabase: storedMappings.length ? storedMappings : materialMappings,
      materialPlans,
      history: storedHistory.length ? storedHistory : history,
      calculation,
      localStorageSnapshot: {
        [PROJECT_STORAGE_KEY]: storedProjects.length ? storedProjects : projects,
        [ACTIVE_PROJECT_STORAGE_KEY]: localStorage.getItem(ACTIVE_PROJECT_STORAGE_KEY) || '',
        [MAPPING_STORAGE_KEY]: storedMappings.length ? storedMappings : materialMappings,
        [HISTORY_STORAGE_KEY]: storedHistory.length ? storedHistory : history,
        [BIM_IMPORTS_STORAGE_KEY]: imports,
        [BIM_ROWS_STORAGE_KEY]: bimRows,
        [MATERIAL_PLANS_STORAGE_KEY]: materialPlans,
        [CALCULATION_STORAGE_KEY]: calculation,
      },
    };

    const blob = new Blob([JSON.stringify(backup, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `bim-carbon-project-backup_${safeFileName(projectName)}_${today()}.json`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
    setStatusMessage('專案資料已匯出為 JSON 備份檔。');
  }

  function restoreBackupData(backup) {
    if (!backup || backup.backupType !== 'bim-carbon-platform-project-backup') {
      throw new Error('invalid backup');
    }

    const nextProjects = Array.isArray(backup.projectDatabase)
      ? backup.projectDatabase.map(normalizeProject).filter((project) => project.project_id)
      : [];
    const restoredActiveProject =
      backup.activeProject && backup.activeProject.project_id
        ? normalizeProject(backup.activeProject)
        : nextProjects.find((project) => project.project_id === backup.activeProjectId) ?? null;
    const nextMappings = Array.isArray(backup.materialMappingDatabase)
      ? backup.materialMappingDatabase.map(normalizeMapping)
      : [];
    const nextHistory = Array.isArray(backup.history) ? backup.history : [];

    localStorage.setItem(PROJECT_STORAGE_KEY, JSON.stringify(nextProjects));
    localStorage.setItem(MAPPING_STORAGE_KEY, JSON.stringify(nextMappings));
    localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(nextHistory));
    localStorage.setItem(BIM_IMPORTS_STORAGE_KEY, JSON.stringify(backup.bimImportSummary ?? []));
    localStorage.setItem(BIM_ROWS_STORAGE_KEY, JSON.stringify(backup.bimRows ?? []));
    localStorage.setItem(MATERIAL_PLANS_STORAGE_KEY, JSON.stringify(backup.materialPlans ?? []));
    localStorage.setItem(CALCULATION_STORAGE_KEY, JSON.stringify(backup.calculation ?? null));
    if (restoredActiveProject?.project_id) {
      localStorage.setItem(ACTIVE_PROJECT_STORAGE_KEY, restoredActiveProject.project_id);
    } else {
      localStorage.removeItem(ACTIVE_PROJECT_STORAGE_KEY);
    }

    setProjects(nextProjects);
    setActiveProject(restoredActiveProject);
    setProjectInfo(restoredActiveProject ?? emptyProjectInfo());
    setImports(Array.isArray(backup.bimImportSummary) ? backup.bimImportSummary : []);
    setBimRows(Array.isArray(backup.bimRows) ? backup.bimRows : []);
    setMaterialMappings(nextMappings);
    setMaterialPlans(Array.isArray(backup.materialPlans) ? backup.materialPlans : []);
    setHistory(restoredActiveProject?.project_id ? filterProjectRecords(nextHistory, restoredActiveProject.project_id) : nextHistory);
    setCalculation(backup.calculation ?? null);
    setViewedFileId('');
    setExpandedPlanIds([]);
    setExpandedMappingFiles([]);
    setIsProjectFormOpen(false);
    setEditingProjectId('');
  }

  function importProjectBackup(file) {
    if (!file) return;
    if (!window.confirm('匯入備份資料將覆蓋目前瀏覽器中的本機資料，是否繼續？')) return;

    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const backup = JSON.parse(String(event.target?.result ?? ''));
        restoreBackupData(backup);
        setStatusMessage('專案資料匯入成功，已完成資料還原。');
        scrollToSection('project-info');
      } catch {
        setStatusMessage('匯入失敗，請確認是否為本平台匯出的專案備份檔。');
      }
    };
    reader.onerror = () => {
      setStatusMessage('匯入失敗，請確認是否為本平台匯出的專案備份檔。');
    };
    reader.readAsText(file);
  }

  function clearLocalPlatformData() {
    if (!window.confirm('確定要清除目前瀏覽器中的本機資料？此操作會移除專案、材料對應、材料方案與歷史分析結果。')) return;

    localStorage.removeItem(PROJECT_STORAGE_KEY);
    localStorage.removeItem(ACTIVE_PROJECT_STORAGE_KEY);
    localStorage.removeItem(MAPPING_STORAGE_KEY);
    localStorage.removeItem(HISTORY_STORAGE_KEY);
    localStorage.removeItem(BIM_IMPORTS_STORAGE_KEY);
    localStorage.removeItem(BIM_ROWS_STORAGE_KEY);
    localStorage.removeItem(MATERIAL_PLANS_STORAGE_KEY);
    localStorage.removeItem(CALCULATION_STORAGE_KEY);
    setProjects([]);
    setActiveProject(null);
    setProjectInfo(emptyProjectInfo());
    setImports([]);
    setBimRows([]);
    setMaterialMappings([]);
    setMaterialPlans([]);
    setHistory([]);
    setCalculation(null);
    setViewedFileId('');
    setExpandedPlanIds([]);
    setExpandedMappingFiles([]);
    setIsProjectFormOpen(false);
    setEditingProjectId('');
    setStatusMessage('已清除本機資料。請建立專案或匯入專案資料備份。');
  }

  const hasActiveProject = Boolean(activeProject?.project_id);

  function filterProjectRecords(records, projectId) {
    return records.filter((record) => record.project_id === projectId);
  }

  function loadProjectScopedData(project) {
    try {
      const savedMappings = JSON.parse(localStorage.getItem(MAPPING_STORAGE_KEY) || '[]');
      const normalizedMappings = Array.isArray(savedMappings) ? savedMappings.map(normalizeMapping) : [];
      setMaterialMappings(normalizedMappings.filter((mapping) => !mapping.project_id || mapping.project_id === project.project_id));
    } catch {
      setMaterialMappings([]);
    }

    try {
      const savedHistory = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || '[]');
      setHistory(Array.isArray(savedHistory) ? filterProjectRecords(savedHistory, project.project_id) : []);
    } catch {
      setHistory([]);
    }

    try {
      const savedImports = JSON.parse(localStorage.getItem(BIM_IMPORTS_STORAGE_KEY) || '[]');
      const savedBimRows = JSON.parse(localStorage.getItem(BIM_ROWS_STORAGE_KEY) || '[]');
      const savedPlans = JSON.parse(localStorage.getItem(MATERIAL_PLANS_STORAGE_KEY) || '[]');
      const savedCalculation = JSON.parse(localStorage.getItem(CALCULATION_STORAGE_KEY) || 'null');
      setImports(Array.isArray(savedImports) ? savedImports.filter((item) => !item.project_id || item.project_id === project.project_id) : []);
      setBimRows(Array.isArray(savedBimRows) ? savedBimRows.filter((row) => !row.project_id || row.project_id === project.project_id) : []);
      setMaterialPlans(Array.isArray(savedPlans) ? savedPlans.filter((plan) => !plan.project_id || plan.project_id === project.project_id) : []);
      setCalculation(savedCalculation?.project_id === project.project_id ? savedCalculation : null);
    } catch {
      clearProjectWorkspace();
    }
  }

  function clearProjectWorkspace() {
    setImports([]);
    setBimRows([]);
    setMaterialPlans([]);
    setCalculation(null);
    setViewedFileId('');
  }

  function requireActiveProject() {
    if (hasActiveProject) return true;
    setStatusMessage('請先建立並儲存專案資訊，才能進行 BIM 資料匯入與碳排分析。');
    scrollToSection('project-info');
    return false;
  }

  function openNewProjectForm() {
    setProjectInfo(emptyProjectInfo());
    setEditingProjectId('');
    setIsProjectFormOpen(true);
    scrollToSection('project-info');
  }

  function openEditProjectForm(project = activeProject) {
    if (!project) {
      openNewProjectForm();
      return;
    }
    const normalizedProject = normalizeProject(project);
    setProjectInfo(normalizedProject);
    setEditingProjectId(normalizedProject.project_id);
    setIsProjectFormOpen(true);
    scrollToSection('project-info');
  }

  function saveProject() {
    const normalizedProject = normalizeProject(projectInfo);
    if (!normalizedProject.projectName.trim() || !normalizedProject.engineeringName.trim()) {
      setStatusMessage('請填寫必填欄位：專案名稱與工程名稱。');
      return;
    }

    const projectToSave = {
      ...normalizedProject,
      project_id: editingProjectId || normalizedProject.project_id || generateProjectId(projects),
    };
    const nextProjects = [
      ...projects.filter((project) => project.project_id !== projectToSave.project_id),
      projectToSave,
    ].sort((a, b) => String(b.createdDate).localeCompare(String(a.createdDate)));

    setProjects(nextProjects);
    setActiveProject(projectToSave);
    setProjectInfo(projectToSave);
    setIsProjectFormOpen(false);
    setEditingProjectId('');
    localStorage.setItem(PROJECT_STORAGE_KEY, JSON.stringify(nextProjects));
    localStorage.setItem(ACTIVE_PROJECT_STORAGE_KEY, projectToSave.project_id);
    clearProjectWorkspace();
    loadProjectScopedData(projectToSave);
    setStatusMessage('專案建立成功，已解鎖 BIM 資料匯入功能。');
  }

  function loadProject(project) {
    const normalizedProject = normalizeProject(project);
    setActiveProject(normalizedProject);
    setProjectInfo(normalizedProject);
    setIsProjectFormOpen(false);
    setEditingProjectId('');
    localStorage.setItem(ACTIVE_PROJECT_STORAGE_KEY, normalizedProject.project_id);
    clearProjectWorkspace();
    loadProjectScopedData(normalizedProject);
    setStatusMessage(`已載入專案 ${normalizedProject.projectName}，請匯入或重新載入該專案 BIM 資料。`);
    scrollToSection('project-info');
  }

  function deleteProject(projectId) {
    if (!window.confirm('確定要刪除此專案？此操作會移除專案清單中的資料。')) return;
    const nextProjects = projects.filter((project) => project.project_id !== projectId);
    setProjects(nextProjects);
    localStorage.setItem(PROJECT_STORAGE_KEY, JSON.stringify(nextProjects));

    if (activeProject?.project_id === projectId) {
      setActiveProject(null);
      setProjectInfo(emptyProjectInfo());
      setIsProjectFormOpen(false);
      setEditingProjectId('');
      clearProjectWorkspace();
      setMaterialMappings([]);
      setHistory([]);
      localStorage.removeItem(ACTIVE_PROJECT_STORAGE_KEY);
      setStatusMessage('已刪除目前專案。請先建立並儲存專案資訊，才能進行 BIM 資料匯入與碳排分析。');
    }
  }

  function saveMappings(nextMappings) {
    setMaterialMappings(nextMappings);
    const savedMappings = JSON.parse(localStorage.getItem(MAPPING_STORAGE_KEY) || '[]');
    const otherMappings = Array.isArray(savedMappings)
      ? savedMappings
          .map(normalizeMapping)
          .filter((mapping) => activeProject?.project_id && mapping.project_id && mapping.project_id !== activeProject.project_id)
      : [];
    localStorage.setItem(MAPPING_STORAGE_KEY, JSON.stringify([...otherMappings, ...nextMappings]));
  }

  function invalidateResults(message) {
    setCalculation(null);
    setStatusMessage(message);
  }

  async function handleBimFileUpload(file, replaceId = '') {
    if (!requireActiveProject()) return;
    if (!file) return;
    const fileId = replaceId || crypto.randomUUID();
    const uploadedAt = new Date().toLocaleString('zh-TW');
    const fileType = file.name.split('.').pop()?.toUpperCase() || 'FILE';

    try {
      const worksheet = await readWorkbook(file);
      const { rows, columns } = worksheetToObjects(worksheet);
      const missingColumns = [
        ['類型', columns.type],
        ['結構材料', columns.material],
        ['體積', columns.volume],
        ['數量', columns.quantity],
      ].filter(([, column]) => !column).map(([label]) => label);
      if (!columns.volume) {
        setImports((current) => [
          ...current.filter((item) => item.id !== fileId),
          {
            id: fileId,
            project_id: activeProject.project_id,
            fileName: file.name,
            fileType,
            uploadedAt,
            status: 'error',
            rowCount: 0,
            totalVolume: 0,
            message: '找不到體積欄位，請確認檔案是否包含「體積」或「Volume」欄位。',
          },
        ]);
        invalidateResults('匯入失敗：找不到體積欄位。');
        return;
      }

      const fileComponentType = getComponentTypeFromFileName(file.name);
      const parsedRows = rows
        .map((row, index) => {
          const type = String(row[columns.type] ?? '').trim();
          const materialName = String(row[columns.material] ?? '').trim();
          const materialCategory = classifyStructuralMaterial(materialName);
          const componentType = fileComponentType;
          const volume = parseNumber(row[columns.volume]);
          const quantity = columns.quantity ? parseNumber(row[columns.quantity]) : 0;
          const componentName = type || componentType;

          return {
            id: `${fileId}-${index}-${Date.now()}`,
            project_id: activeProject.project_id,
            fileId,
            fileName: file.name,
            componentName,
            componentType,
            type,
            materialName,
            materialCategory,
            volume,
            quantity,
            hasVolume: volume > 0,
            hasType: Boolean(type),
            hasMaterial: Boolean(materialName),
            hasQuantity: columns.quantity ? quantity > 0 : false,
          };
        })
        .filter((row) => row.volume > 0);

      setBimRows((current) => [...current.filter((row) => row.fileId !== fileId), ...parsedRows]);
      setImports((current) => [
        ...current.filter((item) => item.id !== fileId),
        {
          id: fileId,
          project_id: activeProject.project_id,
          fileName: file.name,
          fileType,
          uploadedAt,
          status: 'success',
          rowCount: parsedRows.length,
          totalVolume: parsedRows.reduce((sum, row) => sum + row.volume, 0),
          componentType: fileComponentType,
          volumeColumn: columns.volume,
          materialColumn: columns.material || '未提供結構材料欄位',
          typeColumn: columns.type || '未提供類型欄位',
          quantityColumn: columns.quantity || '未提供數量欄位',
          message: missingColumns.length > 0 ? `缺少欄位：${missingColumns.join('、')}` : '欄位檢查完成',
        },
      ]);
      invalidateResults(
        missingColumns.length > 0
          ? `已匯入 ${file.name}，但缺少欄位：${missingColumns.join('、')}。請確認資料後完成材料對應。`
          : `已匯入 ${file.name}，類型、結構材料、體積與數量欄位已取得。`
      );
    } catch (error) {
      setImports((current) => [
        ...current.filter((item) => item.id !== fileId),
        {
          id: fileId,
          project_id: activeProject.project_id,
          fileName: file.name,
          fileType,
          uploadedAt,
          status: 'error',
          rowCount: 0,
          totalVolume: 0,
          message: `檔案讀取失敗：${error.message}`,
        },
      ]);
      invalidateResults(`檔案讀取失敗：${error.message}`);
    }
  }

  function viewImportedFile(fileId) {
    setViewedFileId((current) => (current === fileId ? '' : fileId));
  }

  function removeImportedFile(fileId) {
    if (!window.confirm('確定要刪除此檔案？\n此操作將同步移除相關構件資料與分析結果。')) return;
    setImports((current) => current.filter((item) => item.id !== fileId));
    setBimRows((current) => current.filter((row) => row.fileId !== fileId));
    invalidateResults('已刪除檔案並移除相關構件資料與分析結果。');
  }

  function clearImportedData() {
    if (!requireActiveProject()) return;
    if (!window.confirm('確定要清空全部匯入資料？')) return;
    setImports([]);
    setBimRows([]);
    setMaterialPlans([]);
    setCalculation(null);
    setStatusMessage('已清空 BIM 匯入資料與分析結果，專案資訊與碳排資料庫已保留。');
  }

  function selectDatabaseMaterial(row) {
    if (!requireActiveProject()) return;
    setMappingDraft((current) => ({
      ...current,
      itemName: row.itemName,
      carbonFactor: String(row.coe),
      unit: row.unit,
      source: row.source,
    }));
    scrollToSection('mapping-center');
  }

  function applyRecommendation(row, material) {
    if (!requireActiveProject()) return;
    const mapping = normalizeMapping({
      project_id: activeProject.project_id,
      componentType: row.componentType,
      bimMaterialName: row.bimMaterialName,
      itemName: material.itemName,
      carbonFactor: material.coe,
      unit: material.unit,
      source: material.source,
      status: 'confirmed',
      note: material.reason || '',
    });
    const next = [
      ...materialMappings.filter((item) => normalizeName(item.bimMaterialName) !== normalizeName(row.bimMaterialName)),
      mapping,
    ];
    saveMappings(next);
    invalidateResults('已完成材料對應，請重新開始計算以更新結果。');
  }

  function acceptAllRecommendations() {
    if (!requireActiveProject()) return;
    const recommended = mappingCenterRows.filter((row) => !row.mapping && row.recommendation);
    if (recommended.length === 0) return;
    const next = [...materialMappings];
    recommended.forEach((row) => {
      const mapping = normalizeMapping({
        project_id: activeProject.project_id,
        componentType: row.componentType,
        bimMaterialName: row.bimMaterialName,
        itemName: row.recommendation.itemName,
        carbonFactor: row.recommendation.coe,
        unit: row.recommendation.unit,
        source: row.recommendation.source,
        status: 'confirmed',
        note: row.recommendation.reason,
      });
      const index = next.findIndex((item) => normalizeName(item.bimMaterialName) === normalizeName(mapping.bimMaterialName));
      if (index >= 0) next[index] = mapping;
      else next.push(mapping);
    });
    saveMappings(next);
    invalidateResults(`已接受 ${recommended.length} 筆材料推薦。`);
  }

  function editMapping(row) {
    if (!requireActiveProject()) return;
    const mapping = row.mapping;
    const currentItemName = mapping?.itemName ?? row.recommendation?.itemName ?? '';
    setDatabaseQuery(currentItemName || row.bimMaterialName || '');
    setMappingDraft({
      componentType: row.componentType,
      bimMaterialName: row.bimMaterialName,
      itemName: currentItemName,
      carbonFactor: mapping ? String(mapping.carbonFactor) : row.recommendation ? String(row.recommendation.coe) : '',
      unit: mapping?.unit ?? row.recommendation?.unit ?? 'kgCO2e/m3',
      source: mapping?.source ?? row.recommendation?.source ?? '',
    });
    scrollToSection('mapping-center');
  }

  function saveMappingDraft() {
    if (!requireActiveProject()) return;
    const factor = toNumber(mappingDraft.carbonFactor);
    if (!mappingDraft.bimMaterialName || !mappingDraft.itemName || factor <= 0) {
      setStatusMessage('請完整填寫 BIM材料名稱、標準材料名稱與碳排係數。');
      return;
    }

    const mapping = normalizeMapping({
      ...mappingDraft,
      project_id: activeProject.project_id,
      carbonFactor: factor,
      status: 'confirmed',
    });
    const next = [
      ...materialMappings.filter((item) => normalizeName(item.bimMaterialName) !== normalizeName(mapping.bimMaterialName)),
      mapping,
    ];
    saveMappings(next);
    setMappingDraft({ componentType: '', bimMaterialName: '', itemName: '', carbonFactor: '', unit: 'kgCO2e/m3', source: '' });
    invalidateResults('材料對應已儲存。');
  }

  function createPlan() {
    if (!requireActiveProject()) return;
    const name = `方案${String.fromCharCode(65 + materialPlans.length)}`;
    const assignments = mappingCenterRows.map((row) => {
      const mapping = row.mapping;
      return {
        componentType: row.componentType,
        bimMaterialName: row.bimMaterialName,
        itemName: mapping?.itemName ?? row.recommendation?.itemName ?? '',
        carbonFactor: mapping?.carbonFactor ?? row.recommendation?.coe ?? 0,
        unit: mapping?.unit ?? row.recommendation?.unit ?? 'kgCO2e/m3',
        source: mapping?.source ?? row.recommendation?.source ?? '',
      };
    });
    setMaterialPlans((current) => [...current, { id: crypto.randomUUID(), project_id: activeProject.project_id, name, assignments }]);
    invalidateResults('已新增材料方案，請開始計算以產生結果。');
  }

  function copyPlan(planId) {
    if (!requireActiveProject()) return;
    const plan = materialPlans.find((item) => item.id === planId);
    if (!plan) return;
    setMaterialPlans((current) => [
      ...current,
      { ...plan, id: crypto.randomUUID(), project_id: activeProject.project_id, name: `${plan.name} 複本`, assignments: plan.assignments.map((item) => ({ ...item })) },
    ]);
  }

  function deletePlan(planId) {
    if (!requireActiveProject()) return;
    setMaterialPlans((current) => current.filter((plan) => plan.id !== planId));
    invalidateResults('已刪除材料方案，請重新開始計算。');
  }

  function updatePlanAssignment(planId, assignmentIndex, itemName) {
    if (!requireActiveProject()) return;
    const selected = databaseRows.find((row) => row.itemName === itemName);
    setMaterialPlans((current) =>
      current.map((plan) =>
        plan.id === planId
          ? {
              ...plan,
              assignments: plan.assignments.map((assignment, index) =>
                index === assignmentIndex
                  ? {
                      ...assignment,
                      itemName,
                      carbonFactor: selected?.coe ?? assignment.carbonFactor,
                      unit: selected?.unit ?? assignment.unit,
                      source: selected?.source ?? assignment.source,
                    }
                  : assignment
              ),
            }
          : plan
      )
    );
    invalidateResults('材料方案已修改，請重新開始計算。');
  }

  function ensureDefaultPlan() {
    if (materialPlans.length > 0) return materialPlans;
    return [
      {
        id: 'default-plan',
        project_id: activeProject?.project_id ?? '',
        name: '目前材料對應方案',
        assignments: mappingCenterRows.map((row) => ({
          componentType: row.componentType,
          bimMaterialName: row.bimMaterialName,
          itemName: row.mapping?.itemName ?? '',
          carbonFactor: row.mapping?.carbonFactor ?? 0,
          unit: row.mapping?.unit ?? 'kgCO2e/m3',
          source: row.mapping?.source ?? '',
        })),
      },
    ];
  }

  function resolveAssignmentForRow(plan, row) {
    const key = row.materialName || row.componentType;
    const planAssignment =
      plan.assignments.find((item) => normalizeName(item.bimMaterialName) === normalizeName(key)) ??
      plan.assignments.find((item) => item.componentType === row.componentType);

    if (toNumber(planAssignment?.carbonFactor) > 0) return planAssignment;

    const mapping =
      mappingByKey.get(normalizeName(key)) ??
      mappingByKey.get(normalizeName(row.componentType));

    if (!mapping || toNumber(mapping.carbonFactor) <= 0) return planAssignment;

    return {
      ...planAssignment,
      componentType: row.componentType,
      bimMaterialName: key,
      itemName: mapping.itemName,
      carbonFactor: mapping.carbonFactor,
      unit: mapping.unit,
      source: mapping.source,
    };
  }

  /* eslint-disable no-unreachable */
  function calculatePlan(plan) {
    const componentMapWithFallback = new Map();
    const rowsWithFallback = bimRows.map((row) => {
      const assignment = resolveAssignmentForRow(plan, row);
      const carbonFactor = toNumber(assignment?.carbonFactor);
      const carbon = row.volume * carbonFactor;
      const current = componentMapWithFallback.get(row.componentType) ?? { componentType: row.componentType, volume: 0, carbon: 0, count: 0 };
      current.volume += row.volume;
      current.carbon += carbon;
      current.count += 1;
      componentMapWithFallback.set(row.componentType, current);
      return { ...row, assignment, carbonFactor, carbon };
    });

    return {
      ...plan,
      rows: rowsWithFallback,
      totalVolume: rowsWithFallback.reduce((sum, row) => sum + row.volume, 0),
      totalCarbon: rowsWithFallback.reduce((sum, row) => sum + row.carbon, 0),
      componentResults: [...componentMapWithFallback.values()].sort((a, b) => b.carbon - a.carbon),
    };

    const assignmentByMaterial = new Map(plan.assignments.map((item) => [normalizeName(item.bimMaterialName), item]));
    const componentMap = new Map();

    const rows = bimRows.map((row) => {
      const key = row.materialName || `${row.componentType}（材料待確認）`;
      const assignment = assignmentByMaterial.get(normalizeName(key)) ?? plan.assignments.find((item) => item.componentType === row.componentType);
      const carbonFactor = toNumber(assignment?.carbonFactor);
      const carbon = row.volume * carbonFactor;
      const current = componentMap.get(row.componentType) ?? { componentType: row.componentType, volume: 0, carbon: 0, count: 0 };
      current.volume += row.volume;
      current.carbon += carbon;
      current.count += 1;
      componentMap.set(row.componentType, current);
      return { ...row, assignment, carbonFactor, carbon };
    });

    return {
      ...plan,
      rows,
      totalVolume: rows.reduce((sum, row) => sum + row.volume, 0),
      totalCarbon: rows.reduce((sum, row) => sum + row.carbon, 0),
      componentResults: [...componentMap.values()].sort((a, b) => b.carbon - a.carbon),
    };
  }

  /* eslint-enable no-unreachable */
  function startCalculation() {
    if (!requireActiveProject()) return;
    if (bimRows.length === 0) {
      setStatusMessage('請先匯入 BIM 明細表。');
      return;
    }

    if (mappingStats.completed < mappingStats.total) {
      setStatusMessage('尚有待確認材料，請完成材料對應後再進行碳排分析。');
      scrollToSection('mapping-center');
      return;
    }

    const plans = ensureDefaultPlan();
    const planResults = plans.map(calculatePlan).sort((a, b) => a.totalCarbon - b.totalCarbon);
    if (planResults.every((plan) => plan.totalCarbon <= 0)) {
      setStatusMessage('材料方案缺少有效碳排係數，請回到「材料對應中心」確認碳排係數大於 0，或重新建立材料方案。');
      scrollToSection('mapping-center');
      return;
    }
    const bestPlan = planResults[0];
    const worstPlan = planResults[planResults.length - 1];
    const reduction = worstPlan.totalCarbon - bestPlan.totalCarbon;
    const reductionRate = worstPlan.totalCarbon > 0 ? (reduction / worstPlan.totalCarbon) * 100 : 0;
    const nextCalculation = {
      projectInfo: activeProject,
      project_id: activeProject.project_id,
      calculatedAt: new Date().toLocaleString('zh-TW'),
      planResults,
      bestPlan,
      worstPlan,
      totalVolume,
      reduction,
      reductionRate,
    };

    setCalculation(nextCalculation);
    const analysisDate = new Date().toLocaleString('zh-TW');
    const historyRecord = {
      id: crypto.randomUUID(),
      analysis_id: crypto.randomUUID(),
      analysisDate,
      analysis_date: analysisDate,
      project_id: activeProject.project_id,
      project_name: activeProject.projectName,
      projectName: activeProject.projectName,
      materialPlan: bestPlan.name,
      totalVolume,
      totalCarbon: bestPlan.totalCarbon,
      bestPlan: bestPlan.name,
      snapshot: nextCalculation,
    };
    const savedHistory = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || '[]');
    const allHistory = Array.isArray(savedHistory) ? savedHistory : [];
    const nextAllHistory = [historyRecord, ...allHistory];
    setHistory(filterProjectRecords(nextAllHistory, activeProject.project_id));
    localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(nextAllHistory));
    setStatusMessage('A1-A3 碳排分析已完成，並已自動儲存至歷史分析結果。');
  }

  function deleteHistoryRecord(recordId) {
    const savedHistory = JSON.parse(localStorage.getItem(HISTORY_STORAGE_KEY) || '[]');
    const allHistory = Array.isArray(savedHistory) ? savedHistory : history;
    const nextAllHistory = allHistory.filter((record) => record.id !== recordId);
    setHistory(hasActiveProject ? filterProjectRecords(nextAllHistory, activeProject.project_id) : []);
    localStorage.setItem(HISTORY_STORAGE_KEY, JSON.stringify(nextAllHistory));
  }

  function reloadHistoryRecord(record) {
    if (record.snapshot) {
      setCalculation(record.snapshot);
      setStatusMessage(`已重新載入 ${record.projectName} 的分析結果。`);
      scrollToSection('carbon-analysis');
    }
  }

  function togglePlanDetails(planId) {
    setExpandedPlanIds((current) =>
      current.includes(planId) ? current.filter((id) => id !== planId) : [...current, planId]
    );
  }

  function toggleMappingFile(fileId) {
    setExpandedMappingFiles((current) =>
      current.includes(fileId) ? current.filter((id) => id !== fileId) : [...current, fileId]
    );
  }

  function buildPlanPieData(plan) {
    return {
      labels: plan.componentResults.map((row) => row.componentType),
      datasets: [
        {
          label: '各構件碳排',
          data: plan.componentResults.map((row) => row.carbon),
          backgroundColor: plan.componentResults.map((_, index) => CHART_COLORS[index % CHART_COLORS.length]),
          borderColor: '#ffffff',
          borderWidth: 2,
        },
      ],
    };
  }

  function safeSheetName(value, fallback = '方案') {
    return String(value || fallback).replace(/[\\/?*[\]:]/g, '').slice(0, 31) || fallback;
  }

  function buildComponentDetailRows(planResults = calculation?.planResults ?? []) {
    return planResults.flatMap((plan) => {
      const grouped = new Map();
      plan.rows.forEach((row) => {
        const materialName = row.assignment?.itemName ?? '未對應材料';
        const key = `${plan.name}-${row.componentType}-${materialName}`;
        const current = grouped.get(key) ?? {
          方案名稱: plan.name,
          構件名稱: row.componentType,
          材料名稱: materialName,
          體積: 0,
          碳排係數: row.carbonFactor,
          碳排量: 0,
        };
        current.體積 += row.volume;
        current.碳排量 += row.carbon;
        grouped.set(key, current);
      });
      return [...grouped.values()].map((row) => ({
        ...row,
        體積: Number(row.體積.toFixed(3)),
        碳排係數: Number(toNumber(row.碳排係數).toFixed(4)),
        碳排量: Number(row.碳排量.toFixed(2)),
      }));
    });
  }

  function buildComponentShareRows(plan = calculation?.bestPlan) {
    if (!plan) return [];
    return [
      ...plan.componentResults.map((row) => ({
        方案名稱: plan.name,
        構件名稱: row.componentType,
        體積: Number(row.volume.toFixed(3)),
        碳排量: Number(row.carbon.toFixed(2)),
        占比: plan.totalCarbon > 0 ? `${formatNumber((row.carbon / plan.totalCarbon) * 100, 1)}%` : '0.0%',
      })),
      {
        方案名稱: plan.name,
        構件名稱: '總碳排',
        體積: Number(plan.totalVolume.toFixed(3)),
        碳排量: Number(plan.totalCarbon.toFixed(2)),
        占比: '100.0%',
      },
    ];
  }

  function buildPlanComparisonRows(planResults = calculation?.planResults ?? []) {
    const worstCarbon = Math.max(...planResults.map((plan) => plan.totalCarbon), 0);
    const bestCarbon = Math.min(...planResults.map((plan) => plan.totalCarbon), Infinity);
    return planResults.map((plan) => {
      const reduction = worstCarbon - plan.totalCarbon;
      return {
        方案名稱: plan.name,
        總體積: Number(plan.totalVolume.toFixed(3)),
        總碳排: Number(plan.totalCarbon.toFixed(2)),
        減碳量: Number(reduction.toFixed(2)),
        減碳百分比: worstCarbon > 0 ? `${formatNumber((reduction / worstCarbon) * 100, 1)}%` : '0.0%',
        最佳方案: plan.totalCarbon === bestCarbon ? '是' : '',
      };
    });
  }

  function tableHtml(headers, rows) {
    return `
      <table>
        <thead><tr>${headers.map((header) => `<th>${escapeHtml(header)}</th>`).join('')}</tr></thead>
        <tbody>
          ${rows.map((row) => `<tr>${headers.map((header) => `<td>${escapeHtml(row[header] ?? '')}</td>`).join('')}</tr>`).join('')}
        </tbody>
      </table>
    `;
  }

  function exportExcel() {
    if (!requireActiveProject()) return;
    if (!calculation) return;
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([activeProject]), '專案資訊');
    calculation.planResults.forEach((plan, index) => {
      const summaryRows = [
        { 項目: '專案名稱', 數值: activeProject.projectName },
        { 項目: '方案名稱', 數值: plan.name },
        { 項目: '總體積', 數值: Number(plan.totalVolume.toFixed(3)) },
        { 項目: '總碳排', 數值: Number(plan.totalCarbon.toFixed(2)) },
        { 項目: '材料數量', 數值: new Set(plan.rows.map((row) => row.assignment?.itemName).filter(Boolean)).size },
        { 項目: '構件數量', 數值: plan.componentResults.length },
        {},
        ...buildComponentDetailRows([plan]),
      ];
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summaryRows), safeSheetName(plan.name, `方案${index + 1}`));
    });
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(buildPlanComparisonRows()), '方案比較');
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet([{ 最佳減碳方案: calculation.bestPlan.name, 最高碳排方案: calculation.worstPlan.name, 減碳量: calculation.reduction, 減碳百分比: `${formatNumber(calculation.reductionRate, 1)}%` }]), '最佳減碳方案');
    XLSX.writeFile(workbook, `BIM_A1-A3_${activeProject.project_id}_碳排分析_${today()}.xlsx`);
  }

  function exportPdf() {
    if (!requireActiveProject()) return;
    if (!calculation) return;
    const bestRows = [{ 最佳減碳方案: calculation.bestPlan.name, 最高碳排方案: calculation.worstPlan.name, 減碳量: formatNumber(calculation.reduction), 減碳百分比: `${formatNumber(calculation.reductionRate, 1)}%` }];
    const planComparisonRows = buildPlanComparisonRows();
    const planSections = calculation.planResults
      .map((plan) => {
        const materials = new Set(plan.rows.map((row) => row.assignment?.itemName).filter(Boolean)).size;
        const summaryRows = [
          { 項目: '專案名稱', 數值: activeProject.projectName },
          { 項目: '方案名稱', 數值: plan.name },
          { 項目: '總體積', 數值: `${formatNumber(plan.totalVolume, 3)} m3` },
          { 項目: '總碳排', 數值: `${formatNumber(plan.totalCarbon)} kgCO2e` },
          { 項目: '材料數量', 數值: materials },
          { 項目: '構件數量', 數值: plan.componentResults.length },
        ];
        return `
          <section class="page-section">
            <h2>${escapeHtml(plan.name)} 分析結果</h2>
            ${tableHtml(['項目', '數值'], summaryRows)}
            <h3>構件碳排明細</h3>
            ${tableHtml(['方案名稱', '構件名稱', '材料名稱', '體積', '碳排係數', '碳排量'], buildComponentDetailRows([plan]))}
            <h3>各構件碳排占比</h3>
            ${tableHtml(['方案名稱', '構件名稱', '體積', '碳排量', '占比'], buildComponentShareRows(plan))}
          </section>
        `;
      })
      .join('');
    const reportWindow = window.open('', '_blank', 'width=1100,height=800');
    if (!reportWindow) {
      setStatusMessage('無法開啟 PDF 匯出視窗，請確認瀏覽器允許彈出視窗。');
      return;
    }
    reportWindow.document.write(`
      <!doctype html>
      <html lang="zh-Hant">
        <head>
          <meta charset="utf-8" />
          <title>BIM A1-A3 碳排分析報表</title>
          <style>
            body { color: #1f2933; font-family: "Microsoft JhengHei", Arial, sans-serif; margin: 32px; }
            h1 { font-size: 24px; margin-bottom: 8px; }
            h2 { font-size: 18px; margin: 26px 0 8px; }
            h3 { font-size: 15px; margin: 18px 0 6px; }
            p { color: #667085; line-height: 1.6; }
            table { border-collapse: collapse; margin-top: 16px; width: 100%; }
            th, td { border: 1px solid #d9e2ec; padding: 8px; text-align: right; }
            th:first-child, td:first-child { text-align: left; }
            th { background: #f0f4f8; }
            .cover { display: grid; min-height: 70vh; place-content: center; text-align: center; }
            .cover p { text-align: center; }
            .page-section { break-before: page; page-break-before: always; }
          </style>
        </head>
        <body>
          <section class="cover">
            <h1>BIM A1-A3 碳排分析報表</h1>
            <p>${escapeHtml(activeProject.projectName)}</p>
            <p>${escapeHtml(calculation.calculatedAt)}</p>
          </section>
          <section class="page-section">
            <h2>專案資訊</h2>
            ${tableHtml(['項目', '數值'], [
              { 項目: '專案編號', 數值: activeProject.project_id },
              { 項目: '專案名稱', 數值: activeProject.projectName },
              { 項目: '工程名稱', 數值: activeProject.engineeringName },
              { 項目: '工程地點', 數值: activeProject.location || '未填寫' },
              { 項目: '建築用途', 數值: activeProject.buildingUse || '未填寫' },
              { 項目: '結構形式', 數值: activeProject.structureType || '未填寫' },
            ])}
          </section>
          ${planSections}
          <section class="page-section">
          <h2>方案比較表</h2>
          ${tableHtml(['方案名稱', '總體積', '總碳排', '減碳量', '減碳百分比', '最佳方案'], planComparisonRows)}
          <h2>最佳減碳方案</h2>
          ${tableHtml(['最佳減碳方案', '最高碳排方案', '減碳量', '減碳百分比'], bestRows)}
          </section>
          <script>window.onload = () => { window.focus(); window.print(); };</script>
        </body>
      </html>
    `);
    reportWindow.document.close();
  }

  return (
    <div className="app dashboard-app">
      <aside className="sidebar" aria-label="Dashboard 導覽">
        <div className="sidebar-header">
          <span>BIM A1-A3</span>
          <strong>碳排分析平台</strong>
        </div>
        <nav className="sidebar-nav">
          {NAV_ITEMS.map((item) => (
            <button className={activeSection === item.id ? 'active' : ''} key={item.id} type="button" onClick={() => scrollToSection(item.id)}>
              {item.label}
            </button>
          ))}
        </nav>
      </aside>

      <div className="app-content">
        <header className="hero dashboard-hero">
          <div className="hero-inner">
            <p className="eyebrow">BIM A1-A3 Carbon Dashboard</p>
            <h1>BIM A1-A3 碳排分析平台</h1>
            <p>以 BIM 明細表、碳足跡係數資料庫與材料替代方案為核心，完成 A1-A3 碳排計算、比較與報表匯出。</p>
            <div className="dashboard-overview">
              <article className="metric-card"><span>匯入檔案</span><strong>{imports.filter((item) => item.status === 'success').length}</strong><small>CSV / XLSX</small></article>
              <article className="metric-card"><span>構件資料</span><strong>{bimRows.length}</strong><small>{formatNumber(totalVolume, 3)} m3</small></article>
              <article className="metric-card"><span>材料對應</span><strong>{mappingStats.completed}/{mappingStats.total}</strong><small>{mappingStats.unmapped} 筆待確認</small></article>
              <article className="metric-card total"><span>最佳碳排</span><strong>{calculation ? formatNumber(calculation.bestPlan.totalCarbon) : '-'}</strong><small>kgCO2e</small></article>
            </div>
          </div>
        </header>

        <main className="main-layout">
          <section className="panel scroll-section" id="project-info">
            <div className="section-title">
              <div><h2>專案資訊</h2><p>請先建立並儲存專案，後續 BIM 匯入、材料對應、碳排分析與歷史結果都會關聯到專案編號。</p></div>
              <div className="row-actions">
                <button className="calculate-button" type="button" onClick={openNewProjectForm}>建立專案</button>
                {activeProject && <button className="select-button" type="button" onClick={() => openEditProjectForm()}>編輯專案</button>}
              </div>
            </div>
            <div className="backup-panel">
              <p>目前平台資料儲存在瀏覽器本機端。若需在不同電腦使用，請先匯出專案資料備份，再於另一台電腦匯入。</p>
              <div className="backup-actions">
                <button className="select-button" type="button" onClick={exportProjectBackup}>匯出專案資料</button>
                <label className="select-button upload-inline">
                  匯入專案資料
                  <input
                    type="file"
                    accept=".json,application/json"
                    onChange={(event) => {
                      importProjectBackup(event.target.files?.[0]);
                      event.target.value = '';
                    }}
                  />
                </label>
                <button className="select-button danger" type="button" onClick={clearLocalPlatformData}>清除本機資料</button>
              </div>
            </div>
            {!activeProject && <p className="lock-message">請先建立並儲存專案資訊，才能進行 BIM 資料匯入與碳排分析。</p>}
            {activeProject && (
              <div className="project-summary project-card">
                <span>專案名稱：{activeProject.projectName}</span>
                <span>專案編號：{activeProject.project_id}</span>
                <span>工程名稱：{activeProject.engineeringName}</span>
                <span>工程地點：{activeProject.location || '未填寫'}</span>
                <span>建築用途：{activeProject.buildingUse || '未填寫'}</span>
                <span>結構形式：{activeProject.structureType || '未填寫'}</span>
                <span>樓層數：{activeProject.floors || '未填寫'}</span>
                <span>建立日期：{activeProject.createdDate}</span>
              </div>
            )}
            {isProjectFormOpen && (
              <div className="project-info-panel">
                <div className="project-info-title">
                  <h2>{editingProjectId ? '編輯專案' : '建立專案'}</h2>
                  <button className="select-button" type="button" onClick={() => setIsProjectFormOpen(false)}>收合</button>
                </div>
                <div className="project-info-grid">
                  <label className="plain-field">專案名稱（必填）<input value={projectInfo.projectName} onChange={(event) => updateProject('projectName', event.target.value)} /></label>
                  <label className="plain-field">工程名稱（必填）<input value={projectInfo.engineeringName} onChange={(event) => updateProject('engineeringName', event.target.value)} /></label>
                  <label className="plain-field">工程地點<input value={projectInfo.location} onChange={(event) => updateProject('location', event.target.value)} /></label>
                  <label className="plain-field">建築用途<input value={projectInfo.buildingUse} onChange={(event) => updateProject('buildingUse', event.target.value)} /></label>
                  <label className="plain-field">結構形式<select value={projectInfo.structureType} onChange={(event) => updateProject('structureType', event.target.value)}><option value="RC">RC</option><option value="SRC">SRC</option><option value="SC">SC</option></select></label>
                  <label className="plain-field">樓層數<input type="number" min="0" step="1" value={projectInfo.floors} onChange={(event) => updateProject('floors', event.target.value)} /></label>
                  <label className="plain-field">建築面積<input type="number" min="0" step="0.01" value={projectInfo.buildingArea} onChange={(event) => updateProject('buildingArea', event.target.value)} /></label>
                  <label className="plain-field">建立日期<input type="date" value={projectInfo.createdDate} onChange={(event) => updateProject('createdDate', event.target.value)} /></label>
                  <label className="plain-field note-field">備註<input value={projectInfo.note} onChange={(event) => updateProject('note', event.target.value)} /></label>
                </div>
                <div className="project-form-actions">
                  <button className="calculate-button" type="button" onClick={saveProject}>儲存專案</button>
                </div>
              </div>
            )}
            <div className="table-wrap project-list">
              <table>
                <thead><tr><th>專案編號</th><th>專案名稱</th><th>工程名稱</th><th>建立日期</th><th>操作</th></tr></thead>
                <tbody>
                  {projects.map((project) => (
                    <tr key={project.project_id} className={activeProject?.project_id === project.project_id ? 'selected-row' : ''}>
                      <td>{project.project_id}</td>
                      <td>{project.projectName}</td>
                      <td>{project.engineeringName}</td>
                      <td>{project.createdDate}</td>
                      <td><div className="row-actions"><button className="select-button" type="button" onClick={() => loadProject(project)}>載入專案</button><button className="select-button" type="button" onClick={() => openEditProjectForm(project)}>編輯</button><button className="select-button danger" type="button" onClick={() => deleteProject(project.project_id)}>刪除</button></div></td>
                    </tr>
                  ))}
                  {projects.length === 0 && <tr><td colSpan="5" className="empty-cell">尚未建立專案。</td></tr>}
                </tbody>
              </table>
            </div>
          </section>

          <section className="panel table-panel scroll-section" id="bim-import">
            <div className="section-title">
              <div><h2>BIM資料匯入</h2><p>支援 Revit 明細表 CSV / XLSX，可連續新增、查看、重新上傳、刪除與清空全部。</p></div>
              <div className="row-actions">
                <label className={`upload-button quick-upload-button ${!hasActiveProject ? 'is-disabled' : ''}`}>新增檔案<input type="file" accept=".csv,.xlsx" disabled={!hasActiveProject} onChange={(event) => { handleBimFileUpload(event.target.files?.[0]); event.target.value = ''; }} /></label>
                <button className="select-button danger" type="button" onClick={clearImportedData} disabled={!hasActiveProject || imports.length === 0}>清空全部</button>
              </div>
            </div>
            {!hasActiveProject && <p className="lock-message">請先建立並儲存專案資訊，才能進行 BIM 資料匯入與碳排分析。</p>}
            <div className="table-wrap">
              <table>
                <thead><tr><th>檔案名稱</th><th>檔案類型</th><th>構件數量</th><th>總體積</th><th>上傳時間</th><th>操作</th></tr></thead>
                <tbody>
                  {imports.map((item) => (
                    <tr key={item.id}>
                      <td>{item.fileName}</td>
                      <td>{item.fileType}</td>
                      <td>{item.status === 'success' ? item.rowCount : '-'}</td>
                      <td>{item.status === 'success' ? `${formatNumber(item.totalVolume, 3)} m3` : '-'}</td>
                      <td>{item.uploadedAt}</td>
                      <td>
                        <div className="row-actions">
                          <button className="select-button" type="button" onClick={() => viewImportedFile(item.id)}>查看</button>
                          <label className={`select-button upload-inline ${!hasActiveProject ? 'is-disabled' : ''}`}>重新上傳<input type="file" accept=".csv,.xlsx" disabled={!hasActiveProject} onChange={(event) => { handleBimFileUpload(event.target.files?.[0], item.id); event.target.value = ''; }} /></label>
                          <button className="select-button danger" type="button" onClick={() => removeImportedFile(item.id)} disabled={!hasActiveProject}>刪除</button>
                        </div>
                      </td>
                    </tr>
                  ))}
                  {imports.length === 0 && <tr><td colSpan="6" className="empty-cell">尚未匯入 BIM 明細表。</td></tr>}
                </tbody>
              </table>
            </div>
            {viewedFileId && (
              <div className="table-wrap">
                <table>
                  <thead><tr><th>類型</th><th>構件類別</th><th>結構材料</th><th>體積</th><th>數量</th></tr></thead>
                  <tbody>
                    {bimRows.filter((row) => row.fileId === viewedFileId).slice(0, 80).map((row) => (
                      <tr key={row.id}><td>{row.type || '未提供'}</td><td>{row.componentType}</td><td>{row.materialName || '未提供'}</td><td>{formatNumber(row.volume, 3)}</td><td>{row.quantity || '-'}</td></tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </section>

          <section className="panel table-panel scroll-section" id="bim-check">
            <div className="section-title"><div><h2>BIM資料檢查</h2><p>以固定 Revit 明細表流程檢查欄位：類型、結構材料、體積與數量，接著進行結構材料辨識與材料名稱對應。</p></div></div>
            <div className="quality-grid">
              <article className={quality.missing.type ? 'quality-card warning' : 'quality-card success'}><strong>{quality.missing.type ? '!' : '✓'}</strong><span>類型</span><small>{quality.missing.type ? `${quality.missing.type} 筆待補全` : '已取得'}</small></article>
              <article className={quality.missing.material ? 'quality-card warning' : 'quality-card success'}><strong>{quality.missing.material ? '!' : '✓'}</strong><span>結構材料</span><small>{quality.missing.material ? `${quality.missing.material} 筆待補全` : '已取得'}</small></article>
              <article className={quality.missing.volume ? 'quality-card warning' : 'quality-card success'}><strong>{quality.missing.volume ? '!' : '✓'}</strong><span>體積</span><small>{quality.missing.volume ? `${quality.missing.volume} 筆待補全` : '已取得'}</small></article>
              <article className={quality.missing.quantity ? 'quality-card warning' : 'quality-card success'}><strong>{quality.missing.quantity ? '!' : '✓'}</strong><span>數量</span><small>{quality.missing.quantity ? `${quality.missing.quantity} 筆待補全` : '已取得'}</small></article>
            </div>
            <div className="summary-grid">
              <article className="metric-card"><span>資料完整率</span><strong>{formatNumber(quality.completionRate, 1)}%</strong><small>四項欄位</small></article>
              <article className="metric-card"><span>辨識材料</span><strong>{quality.uniqueMaterialCount}</strong><small>項</small></article>
              <article className="metric-card"><span>材料辨識成功率</span><strong>{formatNumber(mappingStats.recommendationRate, 1)}%</strong><small>{mappingStats.recommended}/{mappingStats.total} 項</small></article>
              <article className="metric-card"><span>材料對應成功率</span><strong>{formatNumber(mappingStats.confirmationRate, 1)}%</strong><small>{mappingStats.completed}/{mappingStats.total} 項</small></article>
              <article className="metric-card"><span>待確認</span><strong>{mappingStats.pending + mappingStats.unmapped}</strong><small>項</small></article>
            </div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>欄位檢查</th><th>待補全筆數</th><th>建議處理方式</th></tr></thead>
                <tbody>
                  {[
                    ['類型', quality.missing.type, '確認明細表含「類型」欄位。'],
                    ['結構材料', quality.missing.material, '確認明細表含「結構材料」欄位，或於材料對應中心手動指定。'],
                    ['體積', quality.missing.volume, '確認明細表含 Volume 或體積欄位。'],
                    ['數量', quality.missing.quantity, '確認明細表含「數量」欄位。'],
                  ].filter((row) => row[1] > 0).map(([field, count, suggestion]) => (
                    <tr key={field}><td>{field}</td><td>{count}</td><td>{suggestion}</td></tr>
                  ))}
                  {quality.missingTotal === 0 && <tr><td colSpan="3" className="empty-cell">{bimRows.length ? '資料已完成補全準備。' : '尚未匯入資料。'}</td></tr>}
                </tbody>
              </table>
            </div>
          </section>

          <section className="panel table-panel scroll-section" id="mapping-center">
            <div className="section-title">
              <div><h2>材料對應中心</h2><p>直接讀取 Revit「結構材料」欄位，辨識材料類別並對應至碳排係數資料庫 item_name。{databaseMessage}</p></div>
              <button className="calculate-button" type="button" onClick={acceptAllRecommendations} disabled={!hasActiveProject || mappingStats.pending === 0}>全部接受推薦</button>
            </div>
            <div className="summary-grid">
              <article className="metric-card"><span>辨識材料</span><strong>{mappingStats.total}</strong><small>項</small></article>
              <article className="metric-card"><span>已完成對應</span><strong>{mappingStats.completed}</strong><small>項</small></article>
              <article className="metric-card"><span>待確認</span><strong>{mappingStats.pending}</strong><small>項</small></article>
              <article className="metric-card"><span>待確認材料</span><strong>{mappingStats.unmapped}</strong><small>項</small></article>
            </div>
            <div className="material-form">
              <label className="plain-field">BIM構件<input value={mappingDraft.componentType} onChange={(event) => setMappingDraft((current) => ({ ...current, componentType: event.target.value }))} /></label>
              <label className="plain-field">結構材料<input value={mappingDraft.bimMaterialName} onChange={(event) => setMappingDraft((current) => ({ ...current, bimMaterialName: event.target.value }))} /></label>
              <label className="plain-field">搜尋標準材料<input value={databaseQuery} onChange={(event) => setDatabaseQuery(event.target.value)} /></label>
              <label className="plain-field">標準材料名稱<select value={mappingDraft.itemName} onChange={(event) => selectDatabaseMaterial(databaseRows.find((row) => row.itemName === event.target.value) ?? { itemName: event.target.value, coe: mappingDraft.carbonFactor, unit: mappingDraft.unit, source: mappingDraft.source })}><option value="">請選擇</option>{mappingSelectRows.map((row) => <option key={row.id} value={row.itemName}>{row.itemName}</option>)}</select></label>
              <label className="plain-field">碳排係數<input type="number" min="0" step="0.0001" value={mappingDraft.carbonFactor} onChange={(event) => setMappingDraft((current) => ({ ...current, carbonFactor: event.target.value }))} /></label>
              <label className="plain-field">單位<input value={mappingDraft.unit} onChange={(event) => setMappingDraft((current) => ({ ...current, unit: event.target.value }))} /></label>
              <label className="plain-field">資料來源<input value={mappingDraft.source} onChange={(event) => setMappingDraft((current) => ({ ...current, source: event.target.value }))} /></label>
              <button className="calculate-button" type="button" onClick={saveMappingDraft} disabled={!hasActiveProject}>儲存對應</button>
            </div>
            <div className="mapping-file-groups">
              {mappingGroups.map((group) => {
                const isExpanded = expandedMappingFiles.includes(group.fileId);
                return (
                  <article className="mapping-file-group" key={group.fileId}>
                    <button className="mapping-group-header" type="button" onClick={() => toggleMappingFile(group.fileId)}>
                      <span>{isExpanded ? '▼' : '▶'} {group.fileName}</span>
                      <small>材料數量：{group.materialCount}項</small>
                      <small>待確認：{group.pending}項</small>
                      <small>已完成：{group.completed}項</small>
                      <small>總體積：{formatNumber(group.totalVolume, 3)} m3</small>
                    </button>
                    {isExpanded && (
                      <div className="table-wrap">
                        <table>
                          <thead><tr><th>檔案來源</th><th>構件類型</th><th>BIM原始材料名稱</th><th>出現筆數</th><th>總體積(m3)</th><th>推薦標準材料</th><th>碳排係數</th><th>辨識狀態</th><th>操作</th></tr></thead>
                          <tbody>
                            {group.rows.map((row) => {
                              const material = row.mapping ?? row.recommendation;
                              return (
                                <tr key={row.id}>
                                  <td>{row.fileName}</td>
                                  <td>{row.componentType}</td>
                                  <td>{row.bimMaterialName}</td>
                                  <td>{row.count}筆</td>
                                  <td>{formatNumber(row.volume, 3)}</td>
                                  <td>{material?.itemName ?? '待手動指定'}<br /><small>{material?.unit ?? '-'}</small></td>
                                  <td>{material ? formatNumber(material.carbonFactor ?? material.coe, 4) : '-'}</td>
                                  <td><span className={`status-pill ${row.status === '已完成對應' ? 'ok' : row.status === '待確認' ? 'pending' : 'warning'}`}>{row.status}</span></td>
                                  <td><div className="row-actions">{row.recommendation && !row.mapping && <button className="select-button" type="button" onClick={() => applyRecommendation(row, row.recommendation)} disabled={!hasActiveProject}>接受推薦</button>}<button className="select-button" type="button" onClick={() => editMapping(row)} disabled={!hasActiveProject}>修改對應</button></div></td>
                                </tr>
                              );
                            })}
                          </tbody>
                        </table>
                      </div>
                    )}
                  </article>
                );
              })}
              {mappingGroups.length === 0 && <p className="empty-panel">匯入 BIM 資料後，材料補全與對應項目會依檔案來源分組顯示於此。</p>}
            </div>
          </section>

          <section className="panel table-panel scroll-section" id="plan-management">
            <div className="section-title">
              <div><h2>材料方案管理</h2><p>建立方案 A、B、C，比較不同混凝土或替代材料對 A1-A3 碳排的影響。</p></div>
              <button className="calculate-button" type="button" onClick={createPlan} disabled={!hasActiveProject || mappingCenterRows.length === 0}>新增方案</button>
            </div>
            <div className="material-filter-bar">
              <label className="plain-field">材料分類
                <select value={planMaterialCategory} onChange={(event) => setPlanMaterialCategory(event.target.value)}>
                  {['全部', '混凝土', '鋼材', '玻璃', '鋁材', '板材', '裝修材料', '其他'].map((category) => (
                    <option key={category} value={category}>{category}</option>
                  ))}
                </select>
              </label>
              <p className="filter-hint">在各方案的材料欄位輸入關鍵字，可即時顯示前 10 筆符合材料，並可用方向鍵與 Enter 選取。</p>
            </div>
            <div className="plan-grid">
              {materialPlans.map((plan) => (
                <article className="panel plan-card" key={plan.id}>
                  <div className="project-info-title">
                    <input value={plan.name} onChange={(event) => setMaterialPlans((current) => current.map((item) => item.id === plan.id ? { ...item, name: event.target.value } : item))} />
                    <div className="row-actions"><button className="select-button" type="button" onClick={() => copyPlan(plan.id)} disabled={!hasActiveProject}>複製</button><button className="select-button danger" type="button" onClick={() => deletePlan(plan.id)} disabled={!hasActiveProject}>刪除</button></div>
                  </div>
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>構件</th><th>材料</th><th>係數</th></tr></thead>
                      <tbody>
                        {plan.assignments.map((assignment, index) => (
                          <tr key={`${plan.id}-${assignment.bimMaterialName}`}>
                            <td>{assignment.componentType}</td>
                            <td>
                              <SearchableMaterialSelector
                                category={planMaterialCategory}
                                databaseRows={databaseRows}
                                disabled={!hasActiveProject}
                                key={`${plan.id}-${assignment.bimMaterialName}-${assignment.itemName}`}
                                onSelect={(row) => updatePlanAssignment(plan.id, index, row.itemName)}
                                value={assignment.itemName}
                              />
                            </td>
                            <td>{formatNumber(assignment.carbonFactor, 4)}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </article>
              ))}
              {materialPlans.length === 0 && <p className="empty-panel">尚未建立材料方案。若直接開始計算，系統會使用目前材料對應作為預設方案。</p>}
            </div>
          </section>

          <section className="panel scroll-section" id="carbon-analysis">
            <div className="section-title">
              <div><h2>A1-A3碳排分析</h2><p>所有材料完成對應後，才允許開始計算。</p></div>
              <button className="calculate-button" type="button" onClick={startCalculation} disabled={!hasActiveProject}>開始計算</button>
            </div>
            <p className="chart-note">{statusMessage}</p>
            {calculation && (
              <>
                <div className="plan-result-grid">
                  {calculation.planResults.map((plan) => {
                    const isExpanded = expandedPlanIds.includes(plan.id);
                    const materialCount = new Set(plan.rows.map((row) => row.assignment?.itemName).filter(Boolean)).size;
                    return (
                      <article className="panel plan-result-card" key={plan.id}>
                        <div className="project-info-title">
                          <div>
                            <h2>{plan.name}</h2>
                            <p>{activeProject.projectName}</p>
                          </div>
                          <button className="select-button" type="button" onClick={() => togglePlanDetails(plan.id)}>
                            {isExpanded ? '收合詳情' : '查看詳情'}
                          </button>
                        </div>
                        <div className="summary-grid plan-summary-grid">
                          <article className="metric-card"><span>總體積</span><strong>{formatNumber(plan.totalVolume, 3)}</strong><small>m3</small></article>
                          <article className="metric-card total"><span>總碳排</span><strong>{formatNumber(plan.totalCarbon)}</strong><small>kgCO2e</small></article>
                          <article className="metric-card"><span>材料數量</span><strong>{materialCount}</strong><small>項</small></article>
                          <article className="metric-card"><span>構件數量</span><strong>{plan.componentResults.length}</strong><small>項</small></article>
                        </div>
                        {isExpanded && (
                          <div className="plan-detail">
                            <div className="chart-box compact-chart"><Pie data={buildPlanPieData(plan)} options={chartOptions} /></div>
                            <div className="table-wrap">
                              <table>
                                <thead><tr><th>構件名稱</th><th>材料名稱</th><th>體積</th><th>碳排係數</th><th>碳排量</th></tr></thead>
                                <tbody>
                                  {buildComponentDetailRows([plan]).map((row) => (
                                    <tr key={`${plan.id}-${row.構件名稱}-${row.材料名稱}`}>
                                      <td>{row.構件名稱}</td><td>{row.材料名稱}</td><td>{formatNumber(row.體積, 3)}</td><td>{formatNumber(row.碳排係數, 4)}</td><td>{formatNumber(row.碳排量)}</td>
                                    </tr>
                                  ))}
                                </tbody>
                              </table>
                            </div>
                          </div>
                        )}
                      </article>
                    );
                  })}
                </div>
                <section className="panel comparison-section">
                  <div className="project-info-title">
                    <h2>方案比較</h2>
                    <span className="status-pill ok">最佳方案：{calculation.bestPlan.name}</span>
                  </div>
                  <div className="chart-box"><Bar data={planBarData} options={chartOptions} /></div>
                  <div className="table-wrap">
                    <table>
                      <thead><tr><th>方案名稱</th><th>總體積</th><th>總碳排</th><th>減碳量</th><th>減碳百分比</th><th>最佳方案</th></tr></thead>
                      <tbody>
                        {buildPlanComparisonRows().map((row) => (
                          <tr key={row.方案名稱}>
                            <td>{row.方案名稱}</td><td>{formatNumber(row.總體積, 3)}</td><td>{formatNumber(row.總碳排)}</td><td>{formatNumber(row.減碳量)}</td><td>{row.減碳百分比}</td><td>{row.最佳方案}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </section>
              </>
            )}
          </section>

          <section className="panel table-panel scroll-section" id="history-results">
            <div className="section-title"><div><h2>歷史分析結果</h2><p>每次開始計算後自動儲存，重新整理頁面後仍會保留。</p></div></div>
            <div className="table-wrap">
              <table>
                <thead><tr><th>分析日期</th><th>專案名稱</th><th>專案編號</th><th>材料方案</th><th>總體積</th><th>總碳排</th><th>最佳方案</th><th>操作</th></tr></thead>
                <tbody>
                  {history.map((record) => (
                    <tr key={record.id}>
                      <td>{record.analysis_date ?? record.analysisDate}</td>
                      <td>{record.project_name ?? record.projectName}</td>
                      <td>{record.project_id ?? '-'}</td>
                      <td>{record.materialPlan}</td>
                      <td>{formatNumber(record.totalVolume, 3)}</td>
                      <td>{formatNumber(record.totalCarbon)}</td>
                      <td>{record.bestPlan}</td>
                      <td><div className="row-actions"><button className="select-button" type="button" onClick={() => reloadHistoryRecord(record)}>查看</button><button className="select-button" type="button" onClick={() => reloadHistoryRecord(record)}>重新載入分析</button><button className="select-button danger" type="button" onClick={() => deleteHistoryRecord(record.id)}>刪除</button></div></td>
                    </tr>
                  ))}
                  {history.length === 0 && <tr><td colSpan="8" className="empty-cell">尚未建立歷史分析結果。</td></tr>}
                </tbody>
              </table>
            </div>
          </section>

          <section className="panel export-panel scroll-section" id="report-export">
            <div className="section-title"><div><h2>報表匯出</h2><p>報表包含專案資訊、材料方案、構件體積、材料對應、碳排結果與圖表摘要。</p></div></div>
            <div className="export-actions"><button className="export-button" type="button" onClick={exportExcel} disabled={!hasActiveProject || !calculation}>匯出Excel</button><button className="export-button secondary" type="button" onClick={exportPdf} disabled={!hasActiveProject || !calculation}>匯出PDF</button></div>
          </section>
        </main>
      </div>
    </div>
  );
}

export default App;
