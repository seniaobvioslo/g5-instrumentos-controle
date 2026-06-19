import React, { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { createRoot } from 'react-dom/client';
import * as XLSX from 'xlsx';
import './styles.css';
import { QUALITY_TABS, buildQualityDashboard } from './issueQualityReportService';

const statusOptions = [
  { value: 'open', label: 'Aberto' },
  { value: 'pending', label: 'Pendente' },
  { value: 'in_progress', label: 'Em andamento' },
  { value: 'in_review', label: 'Em revisão' },
  { value: 'completed', label: 'Concluído' },
  { value: 'closed', label: 'Fechado' }
];

const impactFields = [
  {
    value: 'nivel-impacto',
    label: 'Nível de impacto',
    aliases: ['nivel de impacto', 'nível de impacto', 'impacto previsto', 'impacto previsto ']
  },
  {
    value: 'impacto-escopo',
    label: 'Impacto no escopo',
    aliases: ['impacto no escopo', 'impacto em escopo', 'impacto escopo']
  },
  {
    value: 'impacto-medicao',
    label: 'Impacto em Medição',
    aliases: ['impacto em medicao', 'impacto em medição', 'impacto na medicao', 'impacto medicao']
  },
  {
    value: 'impacto-cronograma',
    label: 'Impacto no Cronograma',
    aliases: ['impacto no cronograma', 'impacto no cronograma:', 'impacto cronograma']
  },
  {
    value: 'fase',
    label: 'Fase',
    aliases: ['fase']
  }
];

const impactLevels = [
  { value: 'Alto', key: 'alto', label: 'Alto', color: '#e94f4f' },
  { value: 'Médio', key: 'medio', label: 'Médio', color: '#f39a12' },
  { value: 'Baixo', key: 'baixo', label: 'Baixo', color: '#4eb567' },
  { value: 'Sem classificação', key: 'sem-classificacao', label: 'Sem classificação', color: '#7b8290' }
];

const impactOptionPalette = ['#0f7c90', '#3f8f70', '#8a6fc7', '#c87536', '#3f6ca8', '#b65d7a', '#6b7f38'];
const eapPreviewLimit = 120;
const UNCLASSIFIED_BUSINESS_UNIT = 'Nao classificado';

const businessUnitOptions = [
  {
    value: 'G5 Engenharia',
    title: 'G5 Engenharia',
    description: 'Gestao e Coordenacao de Projetos BIM'
  },
  {
    value: 'G5 Instrumentos',
    title: 'G5 Instrumentos',
    description: 'Gestao de Obras e Ativos'
  }
];

class AppErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = { error: null };
  }

  static getDerivedStateFromError(error) {
    return { error };
  }

  render() {
    if (this.state.error) {
      return (
        <main className="app-shell">
          <section className="workspace">
            <div className="login-panel">
              <p className="eyebrow">Erro na interface</p>
              <h2>Nao foi possivel exibir esta tela.</h2>
              <p>
                A pagina encontrou um erro inesperado. Recarregue a tela e tente novamente.
                {this.state.error?.message ? ` Detalhe: ${this.state.error.message}` : ''}
              </p>
              <button className="primary-button" type="button" onClick={() => window.location.reload()}>
                Recarregar
              </button>
            </div>
          </section>
        </main>
      );
    }

    return this.props.children;
  }
}

function normalizeText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function formatDate(value) {
  if (!value) return 'Sem prazo';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat('pt-BR', { timeZone: 'UTC' }).format(date);
}

function normalizeBusinessUnit(value) {
  const normalized = normalizeText(value);
  if (!normalized) return UNCLASSIFIED_BUSINESS_UNIT;
  if (normalized.includes('instrument')) return 'G5 Instrumentos';
  if (normalized.includes('engenharia')) return 'G5 Engenharia';
  return UNCLASSIFIED_BUSINESS_UNIT;
}

function getProjectBusinessUnit(project) {
  const candidates = [
    project?.businessUnit,
    project?.business_unit,
    project?.unidadeNegocios,
    project?.unidade_de_negocios,
    project?.['Unidade de negocios'],
    project?.['Unidade de negócios'],
    project?.attributes?.businessUnit,
    project?.attributes?.business_unit,
    project?.attributes?.unidadeNegocios,
    project?.attributes?.['Unidade de negocios'],
    project?.attributes?.['Unidade de negócios'],
    project?.attributes?.extension?.data?.businessUnit,
    project?.attributes?.extension?.data?.unidadeNegocios,
    project?.attributes?.extension?.data?.['Unidade de negocios'],
    project?.attributes?.extension?.data?.['Unidade de negócios']
  ];

  const directValue = candidates.find((candidate) => candidate !== undefined && candidate !== null && String(candidate).trim() !== '');
  return normalizeBusinessUnit(directValue);
}

function sameBusinessUnit(project, businessUnit) {
  return getProjectBusinessUnit(project) === businessUnit;
}

function hasClassifiedBusinessUnit(project) {
  return getProjectBusinessUnit(project) !== UNCLASSIFIED_BUSINESS_UNIT;
}

function formatInputDate(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return '';
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}


function getMonthStart(value = new Date()) {
  const date = value instanceof Date && !Number.isNaN(value.getTime()) ? new Date(value) : new Date();
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

function addMonths(date, amount) {
  const base = getMonthStart(date);
  return new Date(base.getFullYear(), base.getMonth() + amount, 1);
}

function formatPlannerMonthTitle(date) {
  const title = new Intl.DateTimeFormat('pt-BR', { month: 'long', year: 'numeric' }).format(date);
  return title.charAt(0).toUpperCase() + title.slice(1);
}

function buildPlannerMonthDays(monthDate) {
  const monthStart = getMonthStart(monthDate);
  const firstDayOffset = (monthStart.getDay() + 6) % 7;
  const gridStart = addDays(monthStart, -firstDayOffset);
  return Array.from({ length: 42 }, (_, index) => addDays(gridStart, index));
}

const plannerDateFieldAliases = [
  'Data Contratual',
  'Data Prevista G5',
  'Data prevista G5',
  'Data prevista Interna',
  'Data Prevista Interna',
  'Data Entrega Cliente',
  'Data de Entrega Cliente',
  'Data Publicação',
  'Data Publicacao',
  'Due date',
  'Due Date',
  'Prazo',
  'Data de vencimento',
  'Vencimento'
];

function getPlannerIssueDate(issue, customFieldDefinitions = []) {
  const candidates = [
    getScheduleField(issue, customFieldDefinitions, 'dataContratual'),
    getScheduleField(issue, customFieldDefinitions, 'dataMetaInterna'),
    getCustomFieldValueByAliases(issue, customFieldDefinitions, plannerDateFieldAliases),
    issue?.dueDate,
    issue?.due_date,
    issue?.deadline,
    issue?.attributes?.dueDate,
    issue?.attributes?.due_date,
    issue?.attributes?.deadline,
    issue?.raw?.dueDate,
    issue?.raw?.due_date,
    issue?.raw?.attributes?.dueDate,
    issue?.raw?.attributes?.due_date
  ];
  for (const candidate of candidates) {
    const parsed = parseScheduleDate(candidate);
    if (parsed) return parsed;
  }
  return null;
}

function classifyPlannerIssue(issue) {
  const title = String(issue?.title || issue?.name || '');
  const categoryText = normalizeText([
    issue?.category,
    issue?.issueCategory,
    issue?.typeCategory,
    issue?.raw?.category,
    issue?.raw?.attributes?.category
  ].filter(Boolean).join(' '));
  const typeText = normalizeText([
    issue?.issueType,
    issue?.issueSubtype,
    issue?.type,
    issue?.subtype,
    issue?.raw?.issueType,
    issue?.raw?.issueSubtype,
    issue?.raw?.attributes?.issueType,
    issue?.raw?.attributes?.issueSubtype
  ].filter(Boolean).join(' '));

  const isDelivery = categoryText.includes('gestao de entregas') && typeText.includes('marco contratual') && typeText.includes('entrega');
  const isDevelopment = /(^|\s)des\s*-/i.test(title);

  return { isDelivery, isDevelopment };
}

function addDays(date, amount) {
  const next = new Date(date);
  next.setDate(next.getDate() + amount);
  next.setHours(0, 0, 0, 0);
  return next;
}

function isBusinessDay(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return false;
  const day = date.getDay();
  return day !== 0 && day !== 6;
}

function addBusinessDays(date, amount) {
  const next = new Date(date);
  next.setHours(0, 0, 0, 0);
  let remaining = Math.max(0, Number(amount) || 0);
  while (remaining > 0) {
    next.setDate(next.getDate() + 1);
    if (isBusinessDay(next)) remaining -= 1;
  }
  return next;
}

function calculateBusinessDiffDays(startValue, endValue) {
  const startDate = parseIssueDate(startValue);
  const endDate = parseIssueDate(endValue);
  if (!startDate || !endDate) return null;
  if (endDate <= startDate) return 0;
  let count = 0;
  for (let current = addDays(startDate, 1); current <= endDate; current = addDays(current, 1)) {
    if (isBusinessDay(current)) count += 1;
  }
  return count;
}

function listDatesBetween(startDate, endDate) {
  if (!(startDate instanceof Date) || !(endDate instanceof Date)) return [];
  const dates = [];
  for (let current = new Date(startDate); current <= endDate; current = addDays(current, 1)) {
    dates.push(new Date(current));
  }
  return dates;
}

function parseDurationDays(value) {
  const numericValue = Number(String(value || '').replace(',', '.').replace(/[^\d.-]/g, ''));
  return Number.isFinite(numericValue) && numericValue > 0 ? Math.max(1, Math.round(numericValue)) : null;
}

function computeEndDateFromStartAndDays(startValue, daysValue) {
  const startDate = parseIssueDate(startValue);
  const durationDays = parseDurationDays(daysValue);
  if (!startDate || !durationDays) return '';
  return formatInputDate(addBusinessDays(startDate, durationDays));
}

function calculateDateDiffDays(startValue, endValue) {
  const startDate = parseIssueDate(startValue);
  const endDate = parseIssueDate(endValue);
  if (!startDate || !endDate) return null;
  return Math.max(0, Math.ceil((endDate.getTime() - startDate.getTime()) / 86400000));
}

function splitNames(value) {
  if (!value) return [];
  if (Array.isArray(value)) return value.flatMap(splitNames);
  if (typeof value === 'object') {
    const name = value.name || value.displayName || [value.firstName, value.lastName].filter(Boolean).join(' ') || value.email || value.id;
    return name ? [String(name).trim()] : [];
  }
  return String(value)
    .split(/[;,|]/)
    .map((item) => item.trim())
    .filter(Boolean);
}

function parseCurrencyNumber(value) {
  if (value === undefined || value === null || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;

  const normalized = String(value)
    .replace(/\s/g, '')
    .replace(/R\$/gi, '')
    .replace(/\./g, '')
    .replace(',', '.')
    .replace(/[^\d.-]/g, '');
  const number = Number(normalized);
  return Number.isFinite(number) ? number : 0;
}

function formatCurrencyBRL(value) {
  return new Intl.NumberFormat('pt-BR', {
    style: 'currency',
    currency: 'BRL'
  }).format(parseCurrencyNumber(value));
}

function formatDateTime(value) {
  if (!value) return 'Nao informado';
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return String(value);
  return new Intl.DateTimeFormat('pt-BR', {
    dateStyle: 'short',
    timeStyle: 'short',
    timeZone: 'America/Sao_Paulo'
  }).format(date);
}

function getImpactFieldValue(issue, selectedField) {
  const fieldConfig = impactFields.find((field) => field.value === selectedField) || impactFields[0];
  const field = (issue.customAttributes || []).find((attribute) => {
    const normalizedName = normalizeText(attribute.name);
    return fieldConfig.aliases.some((alias) => normalizedName.includes(normalizeText(alias)));
  });

  return formatDetailValue(field?.value);
}

function normalizeImpactLevel(value) {
  const normalizedValue = normalizeText(value);

  if (!value || normalizedValue.includes('nao informado') || normalizedValue.includes('sem classificacao')) return 'Sem classificação';
  if (normalizedValue.includes('alto')) return 'Alto';
  if (normalizedValue.includes('medio')) return 'Médio';
  if (normalizedValue.includes('baixo')) return 'Baixo';
  return value;
}

function normalizeImpactOption(value) {
  const normalizedValue = normalizeText(value);

  if (!value || normalizedValue.includes('nao informado') || normalizedValue.includes('sem classificacao')) return 'Sem classificação';
  return String(value).trim();
}

function formatDetailValue(value) {
  if (value === undefined || value === null || value === '') return 'Nao informado';
  if (typeof value === 'boolean') return value ? 'Sim' : 'Nao';
  if (Array.isArray(value)) return value.map(formatDetailValue).join(', ');
  if (typeof value === 'object') {
    return value.name || value.title || value.displayName || value.value || value.id || JSON.stringify(value);
  }

  return String(value);
}

function emptyIfNotInformed(value) {
  const formatted = formatDetailValue(value);
  return formatted === 'Nao informado' ? '' : formatted;
}

function formatPercentValue(value) {
  const text = emptyIfNotInformed(value);
  if (!text) return '';
  if (String(text).includes('%')) return text;
  const normalizedNumber = Number(String(text).replace(',', '.').replace(/[^\d.-]/g, ''));
  if (Number.isFinite(normalizedNumber)) return `${Math.round(normalizedNumber)}%`;
  return text;
}

function sortRiskLevel(firstRisk, secondRisk) {
  const order = { Alto: 0, 'Médio': 1, Medio: 1, Baixo: 2, 'Sem classificação': 3 };
  return (order[firstRisk] ?? 3) - (order[secondRisk] ?? 3);
}

function getDateInputValue(value) {
  if (!value) return '';
  return new Date(value).toISOString().slice(0, 10);
}

function cleanCustomFieldName(name) {
  return String(name || '').replace(/^CF[\s\-_—:]+/i, '').trim();
}

function normalizeCustomFieldType(type) {
  const normalized = normalizeText(type);
  if (normalized.includes('list') || normalized.includes('option') || normalized.includes('enum') || normalized.includes('pick')) return 'list';
  if (normalized.includes('date')) return 'date';
  if (normalized.includes('number') || normalized.includes('numeric') || normalized.includes('integer') || normalized.includes('decimal')) return 'number';
  return 'text';
}

function isOpenIssue(issue) {
  return !['closed', 'completed'].includes(normalizeText(issue.status));
}

function isOverdue(issue) {
  if (!issue.dueDate || !isOpenIssue(issue)) return false;

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  return new Date(issue.dueDate) < today;
}

function getMonthKey(value) {
  if (!value) return 'sem-prazo';

  const date = new Date(value);
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, '0')}`;
}

function getMonthLabel(monthKey) {
  if (monthKey === 'sem-prazo') return 'Sem prazo';

  const [year, month] = monthKey.split('-').map(Number);
  const date = new Date(Date.UTC(year, month - 1, 1));
  return new Intl.DateTimeFormat('pt-BR', { month: 'short', year: 'numeric', timeZone: 'UTC' }).format(date);
}

function sortByName(firstItem, secondItem) {
  return String(firstItem.name || firstItem.title || '').localeCompare(String(secondItem.name || secondItem.title || ''), 'pt-BR', {
    sensitivity: 'base'
  });
}

function collectReferenceIds(value, ids = new Set()) {
  if (!value) return ids;

  if (typeof value === 'string') {
    ids.add(value);
    return ids;
  }

  if (Array.isArray(value)) {
    value.forEach((item) => collectReferenceIds(item, ids));
    return ids;
  }

  if (typeof value === 'object') {
    const possibleId =
      value.id ||
      value.issueId ||
      value.linkedIssueId ||
      value.referenceId ||
      value.targetId ||
      value.target?.id ||
      value.targetIssue?.id ||
      value.targetEntity?.id ||
      value.data?.id ||
      value.displayId ||
      value.target?.displayId ||
      value.targetIssue?.displayId ||
      value.identifier ||
      value.number ||
      value.code;
    if (possibleId) ids.add(possibleId);

    ['issues', 'linkedIssues', 'relatedIssues', 'references', 'links', 'items', 'target', 'targetIssue', 'targetEntity', 'issue', 'data', 'relationships', 'attributes', 'entities'].forEach((fieldName) => {
      if (value[fieldName]) collectReferenceIds(value[fieldName], ids);
    });
  }

  return ids;
}

function isReferenceFieldName(value) {
  const normalized = normalizeText(value);
  return (
    normalized.includes('referencia') ||
    normalized.includes('vinculo') ||
    normalized.includes('relacion') ||
    normalized.includes('dependenc') ||
    normalized.includes('link')
  );
}

function getIssueReferenceCandidates(issue) {
  const raw = issue.raw || {};
  const ids = new Set();

  [
    raw.linkedIssues,
    raw.relatedIssues,
    raw.references,
    raw.referenceIssues,
    raw.issueLinks,
    raw.links,
    raw.relationships,
    raw.attributes?.linkedIssues,
    raw.attributes?.relatedIssues,
    raw.attributes?.references,
    raw.attributes?.issueLinks
  ].forEach((value) => collectReferenceIds(value, ids));

  (issue.customAttributes || []).forEach((attribute) => {
    if (!isReferenceFieldName(attribute.name) && !isReferenceFieldName(attribute.id)) return;
    collectReferenceIds(attribute.rawValue ?? attribute.value, ids);
  });

  const rawCustomAttributes =
    raw.customAttributes || raw.custom_attributes || raw.customFields || raw.custom_fields || raw.attributes?.customAttributes || {};
  if (rawCustomAttributes && typeof rawCustomAttributes === 'object') {
    Object.entries(rawCustomAttributes).forEach(([key, value]) => {
      const possibleName = value?.name || value?.title || key;
      if (!isReferenceFieldName(possibleName) && !isReferenceFieldName(key)) return;
      collectReferenceIds(value?.value ?? value?.displayValue ?? value, ids);
    });
  }

  ids.delete(issue.id);
  return [...ids];
}

function buildIssueLookupIndexes(issues) {
  const byId = new Map();
  const byDisplayId = new Map();
  const byTitle = new Map();

  issues.forEach((issue) => {
    if (!issue?.id) return;
    byId.set(String(issue.id), issue);
    byId.set(normalizeText(issue.id), issue);
    const displayKey = normalizeText(issue.displayId || '');
    if (displayKey) byDisplayId.set(displayKey, issue);
    const titleKey = normalizeText(issue.title || '');
    if (titleKey) byTitle.set(titleKey, issue);
  });

  return { byId, byDisplayId, byTitle };
}

function resolveIssueReferences(issue, lookup) {
  const candidates = issue.relationshipReferenceIds?.length ? issue.relationshipReferenceIds : getIssueReferenceCandidates(issue);
  const resolved = [];
  const seen = new Set();

  const tryAddIssue = (targetIssue) => {
    if (!targetIssue || !targetIssue.id || seen.has(targetIssue.id) || targetIssue.id === issue.id) return false;
    seen.add(targetIssue.id);
    resolved.push(targetIssue);
    return true;
  };

  candidates.forEach((candidate) => {
    if (candidate === undefined || candidate === null) return;
    const rawCandidate = String(candidate).trim();
    if (!rawCandidate) return;

    if (tryAddIssue(lookup.byId.get(rawCandidate))) return;

    const normalizedCandidate = normalizeText(rawCandidate);
    if (!normalizedCandidate) return;
    if (tryAddIssue(lookup.byDisplayId.get(normalizedCandidate))) return;
    if (tryAddIssue(lookup.byTitle.get(normalizedCandidate))) return;

    const candidateParts = new Set(
      normalizedCandidate
        .split(/[\n,;|]/)
        .map((part) => part.trim())
        .filter(Boolean)
    );

    rawCandidate
      .split(/[\n,;|]/)
      .flatMap((part) => part.match(/[A-Za-z]+-\d+|\d+/g) || [])
      .map((part) => normalizeText(part))
      .filter(Boolean)
      .forEach((part) => candidateParts.add(part));

    candidateParts.forEach((part) => {
      tryAddIssue(lookup.byDisplayId.get(part)) || tryAddIssue(lookup.byTitle.get(part));
    });
  });

  return resolved;
}

function getIssueCustomValue(issue, aliases = []) {
  const normalizedAliases = aliases.map(normalizeFieldKey).filter(Boolean);
  const isBroadAlias = (alias) => alias.length >= 8;
  const hasAliasMatch = (value) => {
    const normalizedValue = normalizeFieldKey(value);
    if (!normalizedValue) return false;
    return normalizedAliases.some((alias) => (
      normalizedValue === alias ||
      (isBroadAlias(alias) && normalizedValue.includes(alias)) ||
      (isBroadAlias(normalizedValue) && alias.includes(normalizedValue))
    ));
  };

  const entries = extractIssueCustomAttributeEntries(issue);
  const matched = entries.find((entry) => {
    const candidates = [
      ...entry.names,
      ...entry.ids,
      entry.meta?.attributeDefinitionId,
      entry.meta?.definitionId,
      entry.meta?.fieldId,
      entry.meta?.fieldName,
      entry.meta?.attributeName,
      entry.meta?.displayName,
      entry.meta?.title,
      entry.meta?.name
    ];
    return candidates.some((candidate) => hasAliasMatch(candidate));
  });

  const rawValue = matched?.value;
  if (rawValue !== undefined && rawValue !== null && rawValue !== '') {
    return toDisplayCustomFieldValue(rawValue) || formatDetailValue(rawValue);
  }

  return '';
}

function parseIssueDate(value) {
  return parseScheduleDate(value);
}

function normalizeFieldKey(value) {
  return normalizeText(String(value || '').replace(/\s|\/|_/g, ''));
}

function getNormalizedFieldCandidates(value) {
  return [value?.name, value?.title, value?.displayName, value?.id, value?.definitionId].filter(Boolean).map(normalizeFieldKey);
}

function toDisplayCustomFieldValue(value) {
  if (value === undefined || value === null || value === '') return '';
  if (Array.isArray(value)) return value.map(toDisplayCustomFieldValue).filter(Boolean).join(', ');
  if (typeof value === 'object') {
    return (
      value.displayValue
      || value.label
      || value.name
      || value.title
      || value.displayName
      || value.email
      || value.value
      || value.id
      || ''
    );
  }
  return String(value).trim();
}

function extractIssueCustomAttributeEntries(issue) {
  const entries = [];
  const pushEntry = (entry, rawValue, fallbackKey = '') => {
    if (!entry && rawValue === undefined) return;
    const source = entry && typeof entry === 'object' ? entry : {};
    const names = [fallbackKey, source.name, source.title, source.displayName, source.label, source.fieldName].filter(Boolean);
    const ids = [source.id, source.definitionId, source.fieldId, source.attributeDefinitionId, source.attributeId].filter(Boolean);
    const value = rawValue ?? source.value ?? source.displayValue ?? source.rawValue ?? source.values ?? source.selectedOption ?? source.option;
    entries.push({ names, ids, value, meta: source });
  };

  const candidates = [
    issue.customAttributes,
    issue.custom_attributes,
    issue.attributes,
    issue.customFields,
    issue.custom_fields,
    issue.custom_attributes_values,
    issue.raw?.customAttributes,
    issue.raw?.custom_attributes,
    issue.raw?.attributes,
    issue.raw?.customFields,
    issue.raw?.custom_fields,
    issue.raw?.custom_attributes_values,
    issue.raw?.attributes?.customAttributes,
    issue.raw?.attributes?.custom_attributes
  ].filter(Boolean);

  candidates.forEach((bucket) => {
    if (Array.isArray(bucket)) {
      bucket.forEach((item, index) => pushEntry(item, item?.value ?? item?.displayValue ?? item?.rawValue, String(index)));
      return;
    }
    if (typeof bucket === 'object') {
      Object.entries(bucket).forEach(([key, value]) => {
        if (value && typeof value === 'object' && !Array.isArray(value)) {
          pushEntry({ ...value, id: value.id || key, name: value.name || key }, value.value ?? value.displayValue ?? value.rawValue, key);
          return;
        }
        pushEntry({ id: key, name: key }, value, key);
      });
    }
  });
  return entries;
}

function getCustomFieldValueByAliases(issue, fieldDefinitions = [], aliases = []) {
  const normalizedAliases = aliases.map(normalizeFieldKey).filter(Boolean);
  const isBroadAlias = (alias) => alias.length >= 8;
  const matchesFieldKey = (value) => {
    const normalizedValue = normalizeFieldKey(value);
    if (!normalizedValue) return false;
    return normalizedAliases.some((alias) => (
      normalizedValue === alias ||
      (isBroadAlias(alias) && normalizedValue.includes(alias)) ||
      (isBroadAlias(normalizedValue) && alias.includes(normalizedValue))
    ));
  };

  const definitionById = new Map();
  const definitionMatches = [];
  (fieldDefinitions || []).forEach((definition) => {
    const ids = [definition.id, definition.attributeDefinitionId, definition.definitionId, definition.key]
      .filter(Boolean)
      .map(String);
    ids.forEach((id) => definitionById.set(id, definition));

    const definitionNames = [
      definition.name,
      definition.title,
      definition.displayName,
      definition.label,
      definition.id,
      definition.attributeDefinitionId,
      definition.definitionId,
      definition.key
    ].filter(Boolean);

    if (definitionNames.some(matchesFieldKey)) {
      definitionMatches.push(definition);
    }
  });

  const definitionIds = new Set(
    definitionMatches.flatMap((definition) => [definition.id, definition.attributeDefinitionId, definition.definitionId, definition.key])
      .filter(Boolean)
      .map(String)
  );
  const optionLookup = new Map();
  definitionMatches.forEach((definition) => {
    (definition.options || definition.allowedValues || definition.values || []).forEach((option) => {
      const label = option.label || option.name || option.title || option.value || option.id || '';
      [option.id, option.valueId, option.key, option.value, option.label, option.name, option.title]
        .filter((item) => item !== undefined && item !== null && item !== '')
        .forEach((item) => optionLookup.set(String(item), label));
    });
  });

  const entries = extractIssueCustomAttributeEntries(issue);
  const matched = entries.find((entry) => {
    const normalizedNames = entry.names.map(normalizeFieldKey);
    const normalizedIds = entry.ids.map((id) => String(id || ''));
    const hasAliasName = normalizedNames.some((name) => matchesFieldKey(name));
    const hasDefinitionId = normalizedIds.some((id) => definitionIds.has(String(id)) || matchesFieldKey(id));
    const hasDefinitionNameViaId = normalizedIds.some((id) => {
      const definition = definitionById.get(String(id));
      if (!definition) return false;
      return [definition.name, definition.title, definition.displayName, definition.label, definition.id]
        .filter(Boolean)
        .some(matchesFieldKey);
    });
    return hasAliasName || hasDefinitionId || hasDefinitionNameViaId;
  });

  if (!matched) return null;
  const rawValue = matched.value;
  if (rawValue === undefined || rawValue === null || rawValue === '') return null;
  if (typeof rawValue === 'object' && !Array.isArray(rawValue)) {
    const optionId = rawValue.optionId || rawValue.valueId || rawValue.id || rawValue.key || rawValue.value;
    if (optionId && optionLookup.has(String(optionId))) return optionLookup.get(String(optionId));
  }
  if (Array.isArray(rawValue)) {
    const value = rawValue.map((item) => {
      if (item && typeof item === 'object') {
        const optionId = item.optionId || item.valueId || item.id || item.key || item.value;
        if (optionId && optionLookup.has(String(optionId))) return optionLookup.get(String(optionId));
      }
      const textValue = toDisplayCustomFieldValue(item);
      if (textValue && optionLookup.has(String(textValue))) return optionLookup.get(String(textValue));
      return textValue;
    }).filter(Boolean).join(', ');
    return value || null;
  }
  const displayValue = toDisplayCustomFieldValue(rawValue);
  if (displayValue && optionLookup.has(String(displayValue))) return optionLookup.get(String(displayValue));
  return displayValue || null;
}

function getCustomDateFieldValue(issue, aliases = []) {
  const raw = getCustomFieldValueByAliases(issue, [], aliases);
  if (!raw) return '';
  const date = parseScheduleDate(raw);
  return date ? formatInputDate(date) : '';
}

function findCustomFieldDefinition(fieldDefinitions = [], aliases = []) {
  const normalizedAliases = aliases.map(normalizeFieldKey).filter(Boolean);
  return (fieldDefinitions || []).find((field) => {
    const names = [field.name, field.title, field.displayName, field.label, field.id, field.definitionId]
      .filter(Boolean)
      .map(normalizeFieldKey);
    return names.some((name) => normalizedAliases.includes(name));
  });
}

function getScheduleField(issue, fieldDefinitions, key) {
  return getCustomFieldValueByAliases(issue, fieldDefinitions, scheduleFieldAliases[key] || [key]) || '';
}

function parseScheduleDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const date = new Date(value);
    date.setHours(0, 0, 0, 0);
    return date;
  }

  const text = String(value).trim();
  const isoDateOnlyMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDateOnlyMatch) {
    const year = Number(isoDateOnlyMatch[1]);
    const month = Number(isoDateOnlyMatch[2]) - 1;
    const day = Number(isoDateOnlyMatch[3]);
    const date = new Date(year, month, day);
    if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
      date.setHours(0, 0, 0, 0);
      return date;
    }
    return null;
  }

  const brazilianMatch = text.match(/^(\d{1,2})[\/.-](\d{1,2})[\/.-](\d{2,4})$/);
  if (brazilianMatch) {
    const day = Number(brazilianMatch[1]);
    const month = Number(brazilianMatch[2]) - 1;
    let year = Number(brazilianMatch[3]);
    if (year < 100) year += year >= 70 ? 1900 : 2000;
    const date = new Date(year, month, day);
    if (date.getFullYear() === year && date.getMonth() === month && date.getDate() === day) {
      date.setHours(0, 0, 0, 0);
      return date;
    }
    return null;
  }

  const date = new Date(text);
  if (Number.isNaN(date.getTime())) return null;
  date.setHours(0, 0, 0, 0);
  return date;
}

function formatBrazilianDate(value) {
  const date = parseScheduleDate(value);
  if (!date) return '';
  return new Intl.DateTimeFormat('pt-BR').format(date);
}

function formatScheduleDateForSave(value) {
  return formatBrazilianDate(value) || String(value || '').trim();
}

function computeScheduleEndDateFromBusinessDays(startValue, daysValue) {
  const startDate = parseScheduleDate(startValue);
  const durationDays = parseDurationDays(daysValue);
  if (!startDate || !durationDays) return '';
  return formatBrazilianDate(addBusinessDays(startDate, durationDays));
}

function getScheduleDateInputValue(value) {
  const date = parseScheduleDate(value);
  return date ? formatInputDate(date) : '';
}

function normalizeScheduleStatus(value) {
  return normalizeText(value).replace(/\s+/g, ' ').trim();
}

function isScheduleCompleted(issue, fieldDefinitions, fields = null) {
  const values = [
    issue?.status,
    fields?.faseFluxo ?? getScheduleField(issue, fieldDefinitions, 'faseFluxo'),
    fields?.statusEntrega ?? getScheduleField(issue, fieldDefinitions, 'statusEntrega'),
    fields?.statusCliente ?? getScheduleField(issue, fieldDefinitions, 'statusCliente')
  ].map(normalizeScheduleStatus);

  if (values.some((value) => schedulePartialStatuses.some((partial) => value.includes(partial)))) return false;
  return values.some((value) => scheduleCompletionStatuses.some((status) => value.includes(status)));
}

function getScheduleProgressFromDeliveryStatus(statusValue) {
  const status = normalizeScheduleStatus(statusValue);
  if (!status) return 0;
  if (['aprovado pelo cliente', 'liberado para medicao', 'encerrado', 'medido'].some((item) => status.includes(item))) return 100;
  if (status.includes('medicao parcial aprovada')) return 75;
  if (status.includes('aprovado com comentarios')) return 80;
  if (status.includes('emitido no gerenciador') || status.includes('emitido ao cliente') || status.includes('em analise pelo cliente')) return 55;
  if (status.includes('em revisao interna') || status.includes('com comentarios')) return 60;
  if (status.includes('em atendimento') || status.includes('em desenvolvimento')) return 35;
  if (status.includes('aguardando emissao') || status.includes('aberto')) return 0;
  return 0;
}

function scheduleImpactLevelFromDelay(days) {
  const numericDays = Number(days || 0);
  if (numericDays <= 0) return 'N.A.';
  if (numericDays <= 3) return 'Baixo';
  if (numericDays <= 10) return 'Médio';
  return 'Alto';
}

function schedulePriorityFromSignals({ delayDays = 0, impact = '', blocked = false, critical = false }) {
  const normalizedImpact = normalizeText(impact);
  if (critical || blocked || delayDays > 10 || normalizedImpact.includes('critico') || normalizedImpact.includes('alto')) return 'Alta';
  if (delayDays > 3 || normalizedImpact.includes('medio')) return 'Média';
  if (delayDays > 0 || normalizedImpact.includes('baixo')) return 'Baixa';
  return 'Baixa';
}

function inferScheduleDeliveryStatus(fields, completed, delayDays) {
  if (completed) return 'Encerrado';
  if (fields.dataAprovacaoFinal) return 'Aprovado pelo cliente';
  if (fields.dataRealRetornoCliente && !fields.dataRealReemissao) return 'Em revisão interna';
  if (fields.dataRealEmissao && !fields.dataRealRetornoCliente) return delayDays > 0 ? 'Sem retorno no prazo' : 'Em análise pelo Cliente';
  return 'Aguardando emissão';
}

function inferScheduleClientStatus(fields, completed, delayDays) {
  if (completed || fields.dataAprovacaoFinal) return 'Aprovada';
  if (fields.dataRealRetornoCliente) return 'Comentada';
  if (fields.dataRealEmissao) return delayDays > 0 ? 'Sem retorno no prazo' : 'Aguardando retorno';
  return 'Não enviada';
}

function inferScheduleAnalysisStatus(fields, delayDays) {
  if (fields.dataAprovacaoFinal) return 'Aprovado';
  if (!fields.dataRealEmissao) return 'Não emitido';
  if (fields.dataRealRetornoCliente) return 'Retornado';
  return delayDays > 0 ? 'Sem retorno no prazo' : 'Em análise';
}

function getBusinessDaysOrDefault(value, fallback) {
  return parseDurationDays(value) || fallback;
}

function calculateScheduleDelay(fields, completed) {
  if (completed) return { days: 0, reference: '', reason: 'Concluído/aprovado' };
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const phase = normalizeScheduleStatus([fields.faseFluxo, fields.statusEntrega, fields.statusCliente].filter(Boolean).join(' '));
  const retornoLimit = parseScheduleDate(fields.dataLimiteRetornoCliente)
    || (parseScheduleDate(fields.dataRealEmissao) ? addBusinessDays(parseScheduleDate(fields.dataRealEmissao), getBusinessDaysOrDefault(fields.prazoAnaliseCliente, 10)) : null);
  const revisaoLimit = parseScheduleDate(fields.dataLimiteRevisaoInterna)
    || (parseScheduleDate(fields.dataRealRetornoCliente) ? addBusinessDays(parseScheduleDate(fields.dataRealRetornoCliente), getBusinessDaysOrDefault(fields.prazoRevisaoInterna, 5)) : null);

  let targetDate = parseScheduleDate(fields.dataLimiteInterna) || parseScheduleDate(fields.terminoPlanejado) || parseScheduleDate(fields.dataMetaInterna);
  let reason = 'Meta interna / término planejado';
  if (fields.dataRealEmissao && !fields.dataRealRetornoCliente) {
    targetDate = retornoLimit;
    reason = 'Retorno do cliente';
  } else if ((fields.dataRealRetornoCliente && !fields.dataRealReemissao) || phase.includes('revisao interna')) {
    targetDate = revisaoLimit;
    reason = 'Revisão interna';
  }

  if (!targetDate || today <= targetDate) return { days: 0, reference: targetDate ? formatBrazilianDate(targetDate) : '', reason };
  return { days: calculateBusinessDiffDays(targetDate, today) || calculateDateDiffDays(targetDate, today) || 0, reference: formatBrazilianDate(targetDate), reason };
}

function getIssueTypeLabel(issue) {
  return [issue?.issueSubtype, issue?.issueType, issue?.type, issue?.category].filter(Boolean).join(' / ') || 'Sem tipo';
}

function isScheduleScopeIssue(issue) {
  const scopeText = normalizeText([
    issue?.category,
    issue?.categoryName,
    issue?.issueType,
    issue?.issueSubtype,
    issue?.type,
    issue?.title,
    issue?.raw?.category,
    issue?.raw?.categoryName,
    issue?.raw?.issueType,
    issue?.raw?.issueSubtype,
    issue?.raw?.attributes?.category,
    issue?.raw?.attributes?.issueType,
    issue?.raw?.attributes?.issueSubtype
  ].filter(Boolean).join(' '));
  return scopeText.includes('gestao de entregas');
}

function normalizePredecessorInput(value) {
  return String(value || '').trim();
}

function parsePredecessorIds(value) {
  const raw = normalizePredecessorInput(value);
  if (!raw) return [];
  return raw.split(',').map((item) => item.trim()).filter(Boolean);
}

function getIssueHumanId(issue) {
  return String(issue?.displayId || issue?.number || issue?.issueNumber || issue?.raw?.displayId || issue?.raw?.number || issue?.id || '').trim();
}

const CODIGO_MARCO_FIELD_ALIASES = [
  'Código do Marco',
  'Codigo do Marco',
  'CÓDIGO DO MARCO',
  'Código Marco',
  'Codigo Marco',
  'Cod. Marco',
  'Cod Marco',
  'Marco Código',
  'Marco Codigo'
];

const PACOTE_MARCO_FIELD_ALIASES = [
  'Pacote/Marco Contratual',
  'Pacote/ Marco Contratual',
  'Pacote / Marco Contratual',
  'Pacote/Código do Marco',
  'Pacote / Código do Marco',
  'Marco Contratual',
  'Código do Marco',
  'Pacote',
  'Marco'
];

const KANBAN_MARCO_TITLE_FIELD_ALIASES = [
  'Marco vinculado',
  'Marco Vinculado',
  'Issue marco',
  'Issue do marco',
  'Issue do Marco',
  'Titulo do Marco',
  'TÃ­tulo do Marco',
  'Marco Contratual',
  'Marco contratual',
  'Pacote / Marco Relacionado',
  'Pacote / marco relacionado',
  'Pacote/Marco Contratual',
  'Pacote / Marco Contratual'
];

const MARCO_CONTRATUAL_FIELD_ALIASES = [
  'Marco Contratual',
  'Marco contratual',
  'Marco / Evento Contratual',
  'Marco',
  'Pacote / Marco Relacionado',
  'Pacote / marco relacionado',
  'Pacote/Marco Contratual',
  'Pacote / Marco Contratual'
];

const scheduleFieldAliases = {
  prioridade: ['Prioridade'],
  marcoContratual: ['Marco Contratual', 'Marco / Evento Contratual', 'Marco'],
  codigoMarco: ['Código do Marco', 'Codigo do Marco', 'CÓDIGO DO MARCO', 'Codigo Marco', 'Código Marco'],
  eapVinculada: ['EAP Vinculada', 'Código EAP', 'Codigo EAP', 'EAP'],
  fase: ['Fase'],
  tipoItemCronograma: ['Tipo de Item do Cronograma'],
  tituloDocumento: ['Título do Documento', 'Titulo do Documento'],
  codigoDocumentoCliente: ['Código do Documento Cliente', 'Codigo do Documento Cliente'],
  codigoDocumentoInterno: ['Código do Documento G5', 'Codigo do Documento G5'],
  codigoAreaWf: ['Código Área / WF', 'Codigo Area / WF', 'Código AWP / WF', 'Codigo AWP / WF'],
  disciplinaEnvolvida: ['Disciplina envolvida', 'Disciplina Envolvida'],
  disciplinaSolicitante: ['Disciplina Solicitante'],
  areaResponsavel: ['Área responsável', 'Area responsavel', 'Área Responsável'],
  areasImpactadas: ['Áreas impactadas', 'Areas impactadas'],
  canalTratativa: ['Canal de tratativa'],
  origemDemanda: ['Origem da demanda'],
  origemComentario: ['Origem do comentário', 'Origem do comentario'],
  pacoteMarcoRelacionado: ['Pacote / marco relacionado', 'Pacote / Marco Relacionado', 'Pacote/Marco Contratual', 'Pacote / Marco Contratual'],
  documentoModeloRelacionado: ['Documento / Modelo relacionado', 'Documento / Modelo Relacionado'],
  frenteObra: ['Frente de Obra'],
  periodoRelatorio: ['Período do Relatório', 'Periodo do Relatorio'],
  dataContratual: ['Data Contratual'],
  dataMetaInterna: ['Data prevista Interna', 'Data Prevista Interna', 'Data Meta Interna', 'Data Prevista G5', 'Data prevista G5'],
  dataPublicacao: ['Data Publicação', 'Data Publicacao'],
  dataEmissao: ['Data Emissão', 'Data Emissao'],
  dataPlanejadaEmissao: ['Data Planeada de Emissão', 'Data Planeada de Emissao', 'Data Planejada de Emissão', 'Data Planejada de Emissao'],
  dataRealEmissao: ['Data Real de Emissão', 'Data Real de Emissao'],
  dataLimiteInterna: ['Data Limite Interna'],
  dataLimiteRetornoCliente: ['Data Limite Retorno Cliente'],
  dataRealRetornoCliente: ['Data Real Retorno Cliente'],
  dataLimiteRevisaoInterna: ['Data Limite Revisão Interna', 'Data Limite Revisao Interna'],
  dataRealReemissao: ['Data Real Reemissão', 'Data Real Reemissao'],
  dataAprovacaoFinal: ['Data Aprovação Final', 'Data Aprovacao Final'],
  inicioPlanejado: ['Início Planejado', 'Inicio Planejado', 'Início Previsto', 'Inicio Previsto', 'Data de Início', 'Data de Inicio', 'Data Inicial', 'Data início planejado', 'Data inicio planejado', 'Start Date', 'Data de início planejada', 'Data de inicio planejada'],
  terminoPlanejado: ['Término Planejado', 'Termino Planejado', 'Término Previsto', 'Termino Previsto', 'Data de Término', 'Data de Termino', 'Data Final', 'Data fim planejado', 'Data de vencimento', 'Data Vencimento', 'Prazo', 'Due Date', 'Data prevista'],
  inicioReal: ['Início Real', 'Inicio Real'],
  terminoReal: ['Término real', 'Termino real', 'Término Real', 'Termino Real'],
  diasPrevistosAtividade: ['Dias previstos para atividade', 'Dias Previstos para Atividade', 'Dias previstos atividade'],
  prazoAnaliseCliente: ['Prazo Análise Cliente — dias úteis', 'Prazo Análise Cliente - dias úteis', 'Prazo Analise Cliente - dias uteis'],
  prazoRevisaoInterna: ['Prazo Revisão Interna - dias úteis', 'Prazo Revisao Interna - dias uteis'],
  faseFluxo: ['Fase do Fluxo'],
  statusEntrega: ['Status da entrega', 'Status da Entrega'],
  statusCliente: ['Status Cliente'],
  statusAnaliseCliente: ['Status da Análise Cliente', 'Status da Analise Cliente'],
  statusAsBuilt: ['Status As Built'],
  numeroTramitacao: ['Nº da Tramitação', 'N° da Tramitação', 'Numero da Tramitação', 'Número da Tramitação'],
  numeroEmissaoCliente: ['Número da emissão para o Cliente', 'Numero da emissao para o Cliente'],
  respostaFormalEmitida: ['Resposta formal emitida?'],
  respostaInternaEnviada: ['Resposta G5 foi enviada?', 'Resposta Interna Enviada?'],
  rmAprovadoMedicao: ['RM Aprovado para medição', 'RM Aprovado para medicao'],
  impactoMarco: ['Impacto no Marco'],
  impactoCronograma: ['Impacto no Cronograma'],
  impactoMedicao: ['Impacto em Medição', 'Impacto em Medicao'],
  impactoPrevisto: ['Impacto Previsto'],
  impactoEscopo: ['Impacto no Escopo'],
  impactoQualidade: ['Impacto na Qualidade'],
  riscoPrevisto: ['Risco Previsto'],
  nivelImpacto: ['Nível de Impacto', 'Nivel de Impacto'],
  tipoRiscoRestricao: ['Tipo de Risco / Restrição', 'Tipo de Risco / Restricao'],
  tipoInterface: ['Tipo de Interface'],
  tipoSitNap: ['Tipo SIT / NAP'],
  numeroSitNap: ['Número SIT / NAP', 'Numero SIT / NAP'],
  motivoPendencia: ['Motivo da Pendência', 'Motivo da Pendencia'],
  causaNaoConformidade: ['Causa/Não conformidade', 'Causa/Nao conformidade'],
  acaoCorretivaNecessaria: ['Ação corretiva necessária?', 'Acao corretiva necessaria?'],
  acaoNecessaria: ['Ação Necessária', 'Acao Necessaria'],
  necessitaReuniao: ['Necessita Reunião?', 'Necessita Reuniao?'],
  encaminhamentoDecisaoFinal: ['Encaminhamento / decisão final', 'Encaminhamento / decisao final'],
  prioridadeGestao: ['Prioridade de Gestão', 'Prioridade de Gestao'],
  valorMedicao: ['Valor da Medição', 'Valor da Medicao'],
  tipoMedicao: ['Tipo de Medição', 'Tipo de Medicao'],
  modeloBimVinculado: ['Modelo BIM Vinculado?'],
  tipoValidacaoBim: ['Tipo de Validação BIM', 'Tipo de Validacao BIM'],
  avaliacaoQualidade: ['Avaliação Qualidade', 'Avaliacao Qualidade'],
  percentualTecnico: ['Percentual Técnico', 'Percentual Tecnico'],
  predecessor: ['Predecessor'],
  dependencia: ['Dependência', 'Dependencia']
};

const scheduleFieldLabels = {
  codigoDocumentoInterno: 'Código do Documento Interno',
  dataMetaInterna: 'Data Meta Interna',
  respostaInternaEnviada: 'Resposta Interna Enviada?',
  dataPlanejadaEmissao: 'Data Planejada de Emissão'
};

const scheduleViewModes = [
  { id: 'executiva', label: 'Visão Executiva' },
  { id: 'coordenacao', label: 'Visão de Coordenação' },
  { id: 'tramitacao', label: 'Visão Tramitação' },
  { id: 'critica', label: 'Visão Crítica' },
  { id: 'completa', label: 'Visão Completa' }
];

const scheduleCompletionStatuses = [
  'aprovado',
  'concluido entrega final',
  'concluido',
  'aprovado pelo cliente',
  'liberado para medicao',
  'encerrado',
  'medido',
  'aprovada',
  'closed',
  'completed',
  'done'
];

const schedulePartialStatuses = [
  'aprovado com comentarios',
  'com comentarios',
  'em atendimento',
  'em revisao interna',
  'em analise pelo cliente',
  'medicao parcial aprovada'
];

const scheduleEditableKeys = [
  'codigoMarco',
  'faseFluxo',
  'statusEntrega',
  'statusCliente',
  'statusAnaliseCliente',
  'dataContratual',
  'dataMetaInterna',
  'dataLimiteInterna',
  'dataPublicacao',
  'inicioPlanejado',
  'inicioReal',
  'terminoReal',
  'diasPrevistosAtividade',
  'dataRealEmissao',
  'dataLimiteRetornoCliente',
  'dataRealRetornoCliente',
  'dataLimiteRevisaoInterna',
  'dataAprovacaoFinal',
  'predecessor',
  'impactoMarco',
  'impactoCronograma',
  'prioridadeGestao',
  'percentualTecnico',
  'acaoNecessaria'
];

function getCustomFieldValue(issue, customFieldDefinitions, fieldName) {
  const aliases =
    fieldName === 'Pacote/Marco Contratual'
      ? PACOTE_MARCO_FIELD_ALIASES
      : fieldName === 'Código do Marco'
        ? CODIGO_MARCO_FIELD_ALIASES
        : fieldName === 'Marco Contratual'
          ? MARCO_CONTRATUAL_FIELD_ALIASES
          : [fieldName];
  const targets = aliases.map((alias) => normalizeFieldKey(alias));
  const definition = (customFieldDefinitions || []).find((field) => {
    const names = [field.name, field.title, field.displayName, field.id].filter(Boolean);
    return names.some((name) => targets.includes(normalizeFieldKey(name)));
  });
  const resolvedAliases = [...aliases];
  if (definition?.id) resolvedAliases.push(definition.id);
  if (definition?.name) resolvedAliases.push(definition.name);
  if (definition?.title) resolvedAliases.push(definition.title);
  return getIssueCustomValue(issue, resolvedAliases);
}

function extractReferencedIssuesFromIssueDetail(issueDetail) {
  const sourceIssueId = issueDetail?.id || issueDetail?.issueId || issueDetail?.raw?.id;
  const sourceDisplayId = issueDetail?.displayId || issueDetail?.raw?.displayId || issueDetail?.number || null;
  const buckets = [
    issueDetail?.references,
    issueDetail?.references?.issues,
    issueDetail?.references?.problems,
    issueDetail?.linkedIssues,
    issueDetail?.linkedItems,
    issueDetail?.linkedEntities,
    issueDetail?.relationships,
    issueDetail?.relationships?.references,
    issueDetail?.relationships?.issues,
    issueDetail?.relationships?.linkedIssues,
    issueDetail?.attributes?.references,
    issueDetail?.attributes?.linkedIssues,
    issueDetail?.included,
    issueDetail?.data?.relationships,
    issueDetail?.data?.relationships?.references,
    issueDetail?.data?.relationships?.linkedIssues,
    issueDetail?.raw?.references,
    issueDetail?.raw?.relationships
  ];
  const links = [];
  const seen = new Set();
  const visit = (value, typeHint = '') => {
    if (!value) return;
    if (Array.isArray(value)) return value.forEach((item) => visit(item, typeHint));
    if (typeof value !== 'object') return;
    const entityType = normalizeText(value.type || value.entityType || value.targetType || typeHint);
    const relationName = normalizeText(value.relationshipType || value.linkType || value.name || typeHint);
    const probablyIssueReference =
      entityType.includes('issue') ||
      entityType.includes('problem') ||
      relationName.includes('issue') ||
      relationName.includes('problem') ||
      normalizeText(typeHint).includes('issue') ||
      normalizeText(typeHint).includes('problem');
    if (!probablyIssueReference) {
      Object.entries(value).forEach(([key, nested]) => {
        if (!['raw', 'meta'].includes(key)) visit(nested, key);
      });
      return;
    }
    const targetIssueId = value.targetIssueId || value.issueId || value.linkedIssueId || value.id || value.targetId || null;
    const targetDisplayId = value.targetDisplayId || value.displayId || value.number || value.identifier || null;
    const targetTitle = value.title || value.summary || value.name || value.targetTitle || null;
    if (targetIssueId || targetDisplayId) {
      const key = `${sourceIssueId || ''}:${targetIssueId || ''}:${targetDisplayId || ''}:${targetTitle || ''}`;
      if (!seen.has(key)) {
        seen.add(key);
        links.push({
          sourceIssueId,
          sourceDisplayId,
          targetIssueId,
          targetDisplayId,
          targetTitle,
          referenceType: value.type || value.relationshipType || typeHint || '',
          raw: value
        });
      }
    }
    Object.entries(value).forEach(([key, nested]) => {
      if (['raw', 'meta'].includes(key)) return;
      visit(nested, key);
    });
  };
  buckets.forEach((bucket) => visit(bucket));
  return links;
}

function getIssueGrouping(issue) {
  const ato = getIssueCustomValue(issue, ['ato']);
  const workPackage = getIssueCustomValue(issue, ['work package', 'work-package', 'wp']);
  if (ato && ato !== 'Nao informado') return { key: `ato:${ato}`, label: `Ato: ${ato}` };
  if (workPackage && workPackage !== 'Nao informado') return { key: `wp:${workPackage}`, label: `Work-package: ${workPackage}` };
  const discipline = getIssueDiscipline(issue);
  return { key: `disc:${discipline}`, label: `Disciplina: ${discipline}` };
}

function getStatusPresentation(status) {
  const normalized = normalizeText(status);
  if (normalized === 'open') return { icon: '🟢', color: '#2f9e44' };
  if (normalized === 'in_progress' || normalized === 'in progress') return { icon: '🟠', color: '#f08c00' };
  if (normalized === 'pending' || normalized === 'in_review') return { icon: '🟡', color: '#d0b000' };
  if (normalized === 'closed' || normalized === 'completed') return { icon: '⚫', color: '#6b7280' };
  if (normalized === 'draft') return { icon: '⚪', color: '#9ca3af' };
  return { icon: '🔵', color: '#3b82f6' };
}

function getIssueStatusLabel(status) {
  const rawStatus = String(status || '').trim();
  if (!rawStatus) return 'Sem status';

  const normalizedStatus = normalizeText(rawStatus).replace(/\s+/g, '_');
  const matchedStatus = statusOptions.find((option) => {
    const optionValue = normalizeText(option.value).replace(/\s+/g, '_');
    const optionLabel = normalizeText(option.label);
    return optionValue === normalizedStatus || optionLabel === normalizeText(rawStatus);
  });

  if (matchedStatus) return matchedStatus.label;
  if (['in_progress', 'inprogress', 'em_andamento', 'em_andamento'].includes(normalizedStatus)) return 'Em andamento';
  if (['in_review', 'em_revisao', 'em_revisão'].includes(normalizedStatus)) return 'Em revisão';
  if (['completed', 'complete', 'done', 'concluido', 'concluído'].includes(normalizedStatus)) return 'Concluído';
  if (['closed', 'fechado', 'encerrado'].includes(normalizedStatus)) return 'Fechado';
  if (['pending', 'pendente'].includes(normalizedStatus)) return 'Pendente';
  if (['open', 'aberto'].includes(normalizedStatus)) return 'Aberto';

  return rawStatus;
}

function getIssueDiscipline(issue) {
  const nativeValue = issue.discipline || issue.raw?.discipline || issue.raw?.attributes?.discipline;
  if (nativeValue) return String(nativeValue);

  const customValue = getIssueCustomValue(issue, ['disciplina', 'discipline', 'area tecnica', 'área técnica']);
  if (customValue && customValue !== 'Nao informado') return customValue;

  return issue.category || issue.issueType || 'Sem disciplina';
}

function getIssueTimelineDate(issue) {
  const dueDate = issue.dueDate ? new Date(issue.dueDate).getTime() : null;
  if (Number.isFinite(dueDate)) return dueDate;
  const customDate = getIssueCustomValue(issue, ['prazo', 'data', 'deadline', 'vencimento']);
  if (customDate) {
    const parsedCustomDate = new Date(customDate).getTime();
    if (Number.isFinite(parsedCustomDate)) return parsedCustomDate;
  }
  const created = issue.createdAt ? new Date(issue.createdAt).getTime() : null;
  if (Number.isFinite(created)) return created;
  const updated = issue.updatedAt ? new Date(issue.updatedAt).getTime() : null;
  if (Number.isFinite(updated)) return updated;
  return 0;
}

function getMilestoneCode(issue, customFieldDefinitions = []) {
  return getCustomFieldValueByAliases(issue, customFieldDefinitions, [
    'Código do Marco',
    'Codigo do Marco',
    'Código Marco',
    'Codigo Marco',
    'Cod. Marco',
    'Cod Marco'
  ]);
}

function getMilestoneLabel(issue, customFieldDefinitions = []) {
  const code = getMilestoneCode(issue, customFieldDefinitions);
  return code || issue.displayId || issue.title || issue.id || 'Sem marco';
}

function isDeliveryMilestoneIssue(issue, customFieldDefinitions = []) {
  const text = normalizeText([
    issue.category,
    issue.issueType,
    issue.issueSubtype,
    issue.title,
    getMilestoneCode(issue, customFieldDefinitions)
  ].filter(Boolean).join(' '));

  return (
    text.includes('marco contratual') ||
    text.includes('marco de entrega')
  );
}

function getIssueReferenceTokens(issue) {
  const tokens = new Set();
  const add = (value) => {
    const normalized = normalizeText(value);
    if (normalized) tokens.add(normalized);
  };

  (issue.relationshipReferenceIds || []).forEach(add);
  getIssueReferenceCandidates(issue).forEach(add);
  return tokens;
}

function issueMatchesReferenceTokens(issue, tokens) {
  if (!issue || !tokens?.size) return false;
  return [issue.id, issue.displayId, issue.title]
    .map((value) => normalizeText(value))
    .filter(Boolean)
    .some((value) => tokens.has(value));
}

function getIssueOpenedTime(issue) {
  const value = issue.openedAt || issue.createdAt || issue.updatedAt || issue.dueDate;
  const time = value ? new Date(value).getTime() : 0;
  return Number.isFinite(time) ? time : 0;
}

function getIssueCalendarTime(issue) {
  const value = issue.dueDate || issue.createdAt || issue.openedAt || issue.updatedAt;
  const time = value ? new Date(value).getTime() : 0;
  return Number.isFinite(time) ? time : 0;
}

function sortIssueHierarchy(firstIssue, secondIssue) {
  const openedDiff = getIssueOpenedTime(secondIssue) - getIssueOpenedTime(firstIssue);
  if (openedDiff !== 0) return openedDiff;

  const calendarDiff = getIssueCalendarTime(firstIssue) - getIssueCalendarTime(secondIssue);
  if (calendarDiff !== 0) return calendarDiff;

  return sortByName(firstIssue, secondIssue);
}

function csvEscape(value) {
  const text = String(value ?? '');
  return `"${text.replace(/"/g, '""')}"`;
}

function htmlEscape(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function formatDocumentFolderLabel(value) {
  return String(value || '');
}

function UiIcon({ name }) {
  const icons = {
    modules: 'M3 3h7v7H3V3Zm11 0h7v7h-7V3ZM3 14h7v7H3v-7Zm11 0h7v7h-7v-7Z',
    export: 'M12 3v12m0 0 4-4m-4 4-4-4M4 17v3h16v-3',
    refresh: 'M20 12a8 8 0 1 1-2.34-5.66M20 4v6h-6',
    search: 'M10.5 18a7.5 7.5 0 1 1 5.3-12.8 7.5 7.5 0 0 1-5.3 12.8Zm5.3-2.2L21 21',
    file: 'M7 3h7l5 5v13H7V3Zm7 0v5h5M10 13h6M10 17h6',
    folder: 'M3 6h7l2 2h9v11H3V6Z',
    check: 'M20 6 9 17l-5-5',
    clock: 'M12 3a9 9 0 1 1 0 18 9 9 0 0 1 0-18Zm0 5v5l3 2',
    close: 'M6 6l12 12M18 6 6 18',
    timeline: 'M7 3v3M17 3v3M4 8h16M5 5h14v16H5V5Zm4 9h6m-6 4h3m7-4-4 4m4 0-4-4',
    request: 'M7 3h7l5 5v13H7V3Zm7 0v5h5M11 13h4m-2-2v4',
    overview: 'M12 12m-3 0a3 3 0 1 0 6 0 3 3 0 1 0-6 0M5 16m-2.5 0a2.5 2.5 0 1 0 5 0 2.5 2.5 0 1 0-5 0M19 6m-2.5 0a2.5 2.5 0 1 0 5 0 2.5 2.5 0 1 0-5 0M7.2 14.8l2.5-1.5m4.5-3.5 2.7-2.1',
    documents: 'M3 7h7l2 2h9v10H3V7Zm0 3h18',
    impact: 'M5 19V9m7 10V5m7 14v-7M3 19h18',
    eap: 'M5 4h14v16H5V4Zm0 5h14M10 4v16m-5-5h14m-5-3 3-3 3 3m-3-3v8',
    interfaces: 'M8 8a3 3 0 1 0 0-6 3 3 0 0 0 0 6Zm8 2a3 3 0 1 0 0-6 3 3 0 0 0 0 6ZM3 21v-2a5 5 0 0 1 10 0v2m1-4a5 5 0 0 1 7 4M17 13h4v4m0-4-5 5',
    bim: 'M12 2 4 6v10l8 4 8-4V6l-8-4Zm0 0v8m-8-4 8 4 8-4M4 16l8-4 8 4M8 8v6m8-6v6',
    project: 'M4 6h16M4 12h16M4 18h16M8 4v4M8 10v4M8 16v4M15 5l2 2 3-4M15 11l2 2 3-4M15 17l2 2 3-4',
    meeting: 'M7 3v3M17 3v3M4 8h16M5 5h14v16H5V5Zm4 7h4m-4 4h7M16 12h.01'
  };

  return (
    <svg aria-hidden="true" className="ui-icon" viewBox="0 0 24 24">
      <path d={icons[name] || icons.file} />
    </svg>
  );
}

const meetingTemplates = [
  {
    id: 'design-review',
    label: 'Design Review',
    title: 'Reuniao de Design Review',
    purpose: 'Revisar modelos, entregaveis e decisoes de projeto antes da publicacao ou do pacote de coordenacao.',
    topics: [
      'Abertura e objetivo da revisao',
      'Modelos e documentos em analise',
      'Issues criticas e pendencias por disciplina',
      'Interferencias, compatibilizacao e decisoes necessarias',
      'Encaminhamentos, responsaveis e prazos'
    ]
  },
  {
    id: 'technical',
    label: 'Tecnica',
    title: 'Reuniao Tecnica',
    purpose: 'Alinhar criterios tecnicos, restricoes, responsabilidades e proximas acoes entre as disciplinas.',
    topics: [
      'Contexto tecnico e pauta principal',
      'Pontos abertos por disciplina',
      'Impactos em prazo, escopo e medicao',
      'Validacoes necessarias e referencias tecnicas',
      'Acoes, responsaveis e data de retorno'
    ]
  },
  {
    id: 'management',
    label: 'Gerencial',
    title: 'Reuniao Gerencial',
    purpose: 'Acompanhar marcos, riscos, avancos, indicadores e decisoes de gestao do projeto.',
    topics: [
      'Resumo executivo do projeto',
      'Marcos contratuais, entregaveis e medicao',
      'Riscos, bloqueios e decisoes pendentes',
      'Indicadores de issues e status das entregas',
      'Plano de acao e compromissos da semana'
    ]
  }
];

const instrumentSectionOptions = [
  { id: 'dashboard', title: 'Dashboard Geral', subtitle: 'Indicadores executivos', icon: 'impact' },
  { id: 'agenda', title: 'Agenda de Equipes', subtitle: 'Matriz e calendario', icon: 'timeline' },
  { id: 'equipamentos', title: 'Controle de Equipamentos', subtitle: 'Inventario e movimentacao', icon: 'bim' },
  { id: 'materiais', title: 'Materiais e Compras', subtitle: 'Lista, compras e ABNT', icon: 'documents' },
  { id: 'obras', title: 'Obras / Projetos', subtitle: 'Visao 360 da obra', icon: 'project' },
  { id: 'cadastros', title: 'Cadastros e Padronizacao', subtitle: 'RH, recursos e catalogos', icon: 'interfaces' },
  { id: 'pendencias', title: 'Pendencias e Alertas', subtitle: 'Criticidade operacional', icon: 'overview' },
  { id: 'relatorios', title: 'Relatorios / Exportacoes', subtitle: 'Preparado para Excel/PDF', icon: 'export' }
];

function App() {
  const [user, setUser] = useState(null);
  const [hubs, setHubs] = useState([]);
  const [projects, setProjects] = useState([]);
  const [issues, setIssues] = useState([]);
  const [issueRelationshipsById, setIssueRelationshipsById] = useState({});
  const [publishedDocuments, setPublishedDocuments] = useState([]);
  const [publishedFolder, setPublishedFolder] = useState(null);
  const [documentListRows, setDocumentListRows] = useState([]);
  const [documentSort, setDocumentSort] = useState({ field: 'discipline', direction: 'asc' });
  const [documentListSpreadsheet, setDocumentListSpreadsheet] = useState(null);
  const [publishedDocumentsPartial, setPublishedDocumentsPartial] = useState(false);
  const [publishedDocumentsMessage, setPublishedDocumentsMessage] = useState('');
  const [documentsUpdatedAt, setDocumentsUpdatedAt] = useState('');
  const [documentsSpreadsheetSaving, setDocumentsSpreadsheetSaving] = useState(false);
  const [documentsSpreadsheetSaveMessage, setDocumentsSpreadsheetSaveMessage] = useState('');
  const [documentSearch, setDocumentSearch] = useState('');
  const [documentStatusFilter, setDocumentStatusFilter] = useState('all');
  const [documentTypeFilter, setDocumentTypeFilter] = useState('all');
  const [selectedHubId, setSelectedHubId] = useState('');
  const [selectedProjectId, setSelectedProjectId] = useState('');
  const [selectedBusinessUnit, setSelectedBusinessUnit] = useState(() => localStorage.getItem('central-g5-business-unit') || '');
  const [selectedIssueId, setSelectedIssueId] = useState('');
  const [selectedIssueDetails, setSelectedIssueDetails] = useState(null);
  const [selectedIssueDetailsLoading, setSelectedIssueDetailsLoading] = useState(false);
  const [selectedIssueDetailsError, setSelectedIssueDetailsError] = useState('');
  const [activeModule, setActiveModule] = useState('');
  const [centralBimMenuOpen, setCentralBimMenuOpen] = useState(false);
  const [plannerStartMonth, setPlannerStartMonth] = useState(() => getMonthStart(new Date()));
  const [selectedPlannerDateKey, setSelectedPlannerDateKey] = useState('');
  const [issueTypeOptions, setIssueTypeOptions] = useState([]);
  const [customFieldDefinitions, setCustomFieldDefinitions] = useState([]);
  const [projectUsers, setProjectUsers] = useState([]);
  const [newInterfaceIssue, setNewInterfaceIssue] = useState({
    title: '',
    dueDate: '',
    description: '',
    issueTypeId: '',
    issueSubtypeId: '',
    assignedTo: '',
    customAttributes: {}
  });
  const [projectPickerOpen, setProjectPickerOpen] = useState(false);
  const [issueFilter, setIssueFilter] = useState('all');
  const [categoryFilter, setCategoryFilter] = useState('all');
  const [typeFilter, setTypeFilter] = useState('all');
  const [linkCategoryFilter, setLinkCategoryFilter] = useState('all');
  const [linkTypeFilter, setLinkTypeFilter] = useState('all');
  const [linkStatusFilter, setLinkStatusFilter] = useState('all');
  const [linkMarcoFilter, setLinkMarcoFilter] = useState('all');
  const [linkResponsibleFilter, setLinkResponsibleFilter] = useState('all');
  const [linkCustomFieldFilter, setLinkCustomFieldFilter] = useState('all');
  const [kanbanDraft, setKanbanDraft] = useState({
    marcoId: '',
    type: 'Solicitação de Informação',
    title: '',
    assignedTo: '',
    dueDate: '',
    disciplina: '',
    description: ''
  });
  const [kanbanColumnDraft, setKanbanColumnDraft] = useState({
    title: '',
    code: ''
  });
  const [kanbanManualColumns, setKanbanManualColumns] = useState([]);
  const [kanbanColumnMessage, setKanbanColumnMessage] = useState('');
  const [kanbanCreating, setKanbanCreating] = useState(false);
  const [kanbanMovingIssueId, setKanbanMovingIssueId] = useState('');
  const [kanbanMoveDrafts, setKanbanMoveDrafts] = useState({});
  const [kanbanSavingMoves, setKanbanSavingMoves] = useState(false);
  const [kanbanDraggedIssueId, setKanbanDraggedIssueId] = useState('');
  const [kanbanDragOverColumnId, setKanbanDragOverColumnId] = useState('');
  const [selectedKanbanIssueId, setSelectedKanbanIssueId] = useState('');
  const [kanbanEditDraft, setKanbanEditDraft] = useState({
    marcoId: '',
    title: '',
    assignedTo: '',
    followers: [],
    dueDate: '',
    disciplina: '',
    description: '',
    comment: ''
  });
  const [kanbanSavingIssueId, setKanbanSavingIssueId] = useState('');
  const [selectedImpactField, setSelectedImpactField] = useState(impactFields[0].value);
  const [impactStartDate, setImpactStartDate] = useState('');
  const [impactEndDate, setImpactEndDate] = useState('');
  const [impactTypeFilter, setImpactTypeFilter] = useState('all');
  const [impactCategoryFilter, setImpactCategoryFilter] = useState('all');
  const [impactPhaseFilter, setImpactPhaseFilter] = useState('all');
  const [interfacesDisciplineFilter, setInterfacesDisciplineFilter] = useState('all');
  const [interfacesStatusFilter, setInterfacesStatusFilter] = useState('all');
  const [interfacesImpactFilter, setInterfacesImpactFilter] = useState('all');
  const [loading, setLoading] = useState(false);
  const [projectsLoading, setProjectsLoading] = useState(false);
  const [issuesLoading, setIssuesLoading] = useState(false);
  const [documentsLoading, setDocumentsLoading] = useState(false);
  const [eapFileName, setEapFileName] = useState('');
  const [eapPreview, setEapPreview] = useState(null);
  const [eapResults, setEapResults] = useState([]);
  const [eapExecutionHistory, setEapExecutionHistory] = useState([]);
  const [eapLoading, setEapLoading] = useState(false);
  const [eapDryRun, setEapDryRun] = useState(true);
  const [eapEditedRows, setEapEditedRows] = useState([]);
  const [eapEditedCells, setEapEditedCells] = useState({});
  const [imMode, setImMode] = useState('');
  const [imSheetRows, setImSheetRows] = useState([]);
  const [imSheetOriginalRows, setImSheetOriginalRows] = useState([]);
  const [imSheetCustomColumns, setImSheetCustomColumns] = useState([]);
  const [imSheetSelectedRows, setImSheetSelectedRows] = useState({});
  const [imSheetChanges, setImSheetChanges] = useState({});
  const [imSheetPreviewRows, setImSheetPreviewRows] = useState([]);
  const [imFileName, setImFileName] = useState('');
  const [imPreview, setImPreview] = useState(null);
  const [imResults, setImResults] = useState([]);
  const [imApplySummary, setImApplySummary] = useState(null);
  const [selectedMeetingTemplateId, setSelectedMeetingTemplateId] = useState(meetingTemplates[0].id);
  const [meetingDate, setMeetingDate] = useState('');
  const [meetingParticipants, setMeetingParticipants] = useState('');
  const [meetingCustomTopics, setMeetingCustomTopics] = useState(meetingTemplates[0].topics.join('\n'));
  const [imActionFeedback, setImActionFeedback] = useState(null);
  const [imWorksheetFeedback, setImWorksheetFeedback] = useState(null);
  const [imLoading, setImLoading] = useState(false);
  const [imAllowClearEmpty, setImAllowClearEmpty] = useState(false);
  const [savingIssueId, setSavingIssueId] = useState('');
  const [cronogramaEdits, setCronogramaEdits] = useState({});
  const [scheduleEdits, setScheduleEdits] = useState({});
  const [scheduleSelectedRowId, setScheduleSelectedRowId] = useState('');
  const [schedulePlannerLayout, setSchedulePlannerLayout] = useState({});
  const [scheduleRowHighlights, setScheduleRowHighlights] = useState({});
  const [scheduleNewCategory, setScheduleNewCategory] = useState('Gestão de Entregas');
  const [scheduleViewMode, setScheduleViewMode] = useState('executiva');
  const [scheduleGroupBy, setScheduleGroupBy] = useState('codigoMarco');
  const [scheduleSearch, setScheduleSearch] = useState('');
  const [scheduleStatusFilter, setScheduleStatusFilter] = useState('all');
  const [scheduleOwnerFilter, setScheduleOwnerFilter] = useState('all');
  const [scheduleMarcoFilter, setScheduleMarcoFilter] = useState('all');
  const [scheduleAdvancedFiltersOpen, setScheduleAdvancedFiltersOpen] = useState(false);
  const [scheduleFormulaSelections, setScheduleFormulaSelections] = useState({});
  const [savingSchedule, setSavingSchedule] = useState(false);
  const [scheduleAutosaveState, setScheduleAutosaveState] = useState('idle');
  const [scheduleAutosaveMessage, setScheduleAutosaveMessage] = useState('');
  const [qualityTab, setQualityTab] = useState(QUALITY_TABS[0]);
  const [qualityData, setQualityData] = useState(null);
  const [qualityReportInfo, setQualityReportInfo] = useState(null);
  const [savingCronograma, setSavingCronograma] = useState(false);
  const [issueReportInfo, setIssueReportInfo] = useState(null);
  const [issueReportStatus, setIssueReportStatus] = useState('');
  const [issueReportLoading, setIssueReportLoading] = useState(false);
  const [cronogramaCurrentPage, setCronogramaCurrentPage] = useState(1);
  const [cronogramaDashboardFilter, setCronogramaDashboardFilter] = useState(null);
  const [cronogramaSort, setCronogramaSort] = useState({ key: 'eapVinculada', direction: 'asc' });
  const [eapPanelMode, setEapPanelMode] = useState('dashboard');
  const [eapStructureView, setEapStructureView] = useState('entregas');
  const [eapQuickFilter, setEapQuickFilter] = useState('all');
  const [eapSelectedDeliveryId, setEapSelectedDeliveryId] = useState('all');
  const [eapSelectedRowId, setEapSelectedRowId] = useState('');
  const [projectManagementSearch, setProjectManagementSearch] = useState('');
  const [projectManagementStatusFilter, setProjectManagementStatusFilter] = useState('all');
  const [bimParameterDraft, setBimParameterDraft] = useState({
    name: '',
    type: 'Texto',
    target: 'RVT e DWG',
    value: '',
    notes: ''
  });
  const [bimParameterPlan, setBimParameterPlan] = useState([]);
  const [bimParameterMessage, setBimParameterMessage] = useState('');
  const [instrumentViewMode, setInstrumentViewMode] = useState('quarter');
  const [instrumentStartDate, setInstrumentStartDate] = useState('');
  const [instrumentEndDate, setInstrumentEndDate] = useState('');
  const [instrumentWorkFilter, setInstrumentWorkFilter] = useState('all');
  const [instrumentResourceFilter, setInstrumentResourceFilter] = useState('all');
  const [instrumentFunctionFilter, setInstrumentFunctionFilter] = useState('all');
  const [instrumentStatusFilter, setInstrumentStatusFilter] = useState('all');
  const [instrumentLocationFilter, setInstrumentLocationFilter] = useState('all');
  const [instrumentResourceTypeFilter, setInstrumentResourceTypeFilter] = useState('all');
  const [instrumentMaterialFilter, setInstrumentMaterialFilter] = useState('all');
  const [activeInstrumentSection, setActiveInstrumentSection] = useState('dashboard');
  const [selectedInstrumentAllocationId, setSelectedInstrumentAllocationId] = useState('');
  const [instrumentResourceOverrides, setInstrumentResourceOverrides] = useState({});
  const [error, setError] = useState('');
  const documentTopScrollRef = useRef(null);
  const documentTableScrollRef = useRef(null);
  const scheduleAutosaveTimerRef = useRef(null);

  const loginMessage = useMemo(() => {
    const params = new URLSearchParams(window.location.search);
    if (params.get('login') === 'error') {
      return params.get('message') || 'Falha ao autenticar com Autodesk.';
    }
    return '';
  }, []);

  const selectedProject = useMemo(
    () => projects.find((project) => project.id === selectedProjectId),
    [projects, selectedProjectId]
  );

  const selectedMeetingTemplate = useMemo(
    () => meetingTemplates.find((item) => item.id === selectedMeetingTemplateId) || meetingTemplates[0],
    [selectedMeetingTemplateId]
  );

  const selectedHub = useMemo(
    () => hubs.find((hub) => hub.id === selectedHubId),
    [hubs, selectedHubId]
  );

  const selectedBusinessUnitInfo = useMemo(
    () => businessUnitOptions.find((unit) => unit.value === selectedBusinessUnit) || null,
    [selectedBusinessUnit]
  );

  const filteredProjects = useMemo(() => {
    if (!selectedBusinessUnit) return [];
    return projects;
  }, [projects, selectedBusinessUnit]);

  const issueCategories = useMemo(
    () =>
      issueTypeOptions
        .filter((item) => item.kind === 'type')
        .sort((firstItem, secondItem) => String(firstItem.title).localeCompare(String(secondItem.title), 'pt-BR', { sensitivity: 'base' })),
    [issueTypeOptions]
  );

  const issueSubtypes = useMemo(
    () =>
      issueTypeOptions
        .filter((item) => item.kind === 'subtype' && item.typeId === newInterfaceIssue.issueTypeId)
        .sort((firstItem, secondItem) => String(firstItem.title).localeCompare(String(secondItem.title), 'pt-BR', { sensitivity: 'base' })),
    [issueTypeOptions, newInterfaceIssue.issueTypeId]
  );

  const plannerEventsByDate = useMemo(() => {
    const eventsByDate = new Map();
    issues.forEach((issue) => {
      const issueDate = getPlannerIssueDate(issue, customFieldDefinitions);
      if (!issueDate) return;
      const classification = classifyPlannerIssue(issue);
      if (!classification.isDelivery && !classification.isDevelopment) return;

      const dateKey = formatInputDate(issueDate);
      const baseEvent = {
        issue,
        date: issueDate,
        dateKey,
        title: issue.title || `Issue ${issue.displayId || issue.id || ''}`.trim(),
        category: issue.category || 'Sem categoria',
        type: issue.issueType || issue.issueSubtype || 'Sem tipo',
        status: issue.status || 'Sem status'
      };
      const existing = eventsByDate.get(dateKey) || [];
      if (classification.isDelivery) existing.push({ ...baseEvent, kind: 'delivery', label: 'Marco Contratual / Entrega' });
      if (classification.isDevelopment) existing.push({ ...baseEvent, kind: 'development', label: 'Desenvolvimento' });
      eventsByDate.set(dateKey, existing);
    });
    return eventsByDate;
  }, [issues, customFieldDefinitions]);

  const plannerVisibleMonths = useMemo(
    () => [0, 1, 2].map((offset) => addMonths(plannerStartMonth, offset)),
    [plannerStartMonth]
  );

  const selectedPlannerEvents = useMemo(() => {
    if (!selectedPlannerDateKey) return [];
    return plannerEventsByDate.get(selectedPlannerDateKey) || [];
  }, [plannerEventsByDate, selectedPlannerDateKey]);

  const plannerEventTotals = useMemo(() => {
    let deliveries = 0;
    let developments = 0;
    plannerEventsByDate.forEach((dayEvents) => {
      dayEvents.forEach((event) => {
        if (event.kind === 'delivery') deliveries += 1;
        if (event.kind === 'development') developments += 1;
      });
    });
    return { deliveries, developments };
  }, [plannerEventsByDate]);

  const openPlannerIssueInKanban = useCallback((issueId) => {
    if (!issueId) return;
    setSelectedKanbanIssueId(issueId);
    setActiveModule('links');
  }, []);

  const instrumentationIssues = useMemo(() => {
    return issues.filter((issue) => {
      const haystack = normalizeText([
        issue.issueType,
        issue.issueSubtype,
        issue.category,
        issue.title,
        getCustomFieldValueByAliases(issue, customFieldDefinitions, ['Tipo de issue', 'Tipo', 'Tipo do Issue'])
      ].filter(Boolean).join(' '));
      return haystack.includes('controle de obras') && haystack.includes('instrument');
    });
  }, [issues, customFieldDefinitions]);

  const projectUserLookup = useMemo(() => {
    const aliases = new Map();
    const users = new Map();

    projectUsers.forEach((projectUser) => {
      const key = normalizeText(projectUser.name || projectUser.email || projectUser.id);
      if (!key) return;
      const normalizedUser = {
        key,
        name: projectUser.name || projectUser.email || projectUser.id,
        email: projectUser.email || '',
        id: projectUser.id || '',
        company: projectUser.company || projectUser.companyName || '',
        operationalRole: projectUser.role || projectUser.projectRole || ''
      };
      users.set(key, normalizedUser);
      [projectUser.id, projectUser.name, projectUser.email, projectUser.autodeskId, projectUser.userId, projectUser.uid]
        .filter(Boolean)
        .forEach((value) => aliases.set(normalizeText(value), normalizedUser));
    });

    return { aliases, users };
  }, [projectUsers]);

  const resolveKnownProjectUser = useCallback((value) => {
    if (!value) return null;
    if (Array.isArray(value)) return value.map(resolveKnownProjectUser).find(Boolean) || null;

    const directValue =
      typeof value === 'object'
        ? value.name || value.displayName || [value.firstName, value.lastName].filter(Boolean).join(' ') || value.email || value.id || value.userId || value.autodeskId || value.uid
        : value;
    const normalized = normalizeText(directValue);
    if (!normalized) return null;

    const matchedUser = projectUserLookup.aliases.get(normalized);
    if (matchedUser) return matchedUser;

    const rawText = String(directValue).trim();
    const looksLikeAutodeskId = /^[a-z0-9]{10,}$/i.test(rawText) && !rawText.includes(' ');
    if (looksLikeAutodeskId) return null;

    return { key: normalized, name: rawText, email: '', id: '', company: '', operationalRole: '' };
  }, [projectUserLookup]);

  const instrumentationAllocations = useMemo(() => {
    const readField = (issue, aliases) => getCustomFieldValueByAliases(issue, customFieldDefinitions, aliases) || getIssueCustomValue(issue, aliases) || '';

    return instrumentationIssues.flatMap((issue) => {
      const startDate = parseIssueDate(
        readField(issue, ['Data de inicio', 'Data de início', 'Inicio', 'Start Date']) ||
          issue.startDate ||
          issue.raw?.startDate ||
          issue.raw?.attributes?.startDate
      );
      const endDate =
        parseIssueDate(
          readField(issue, ['Data de vencimento', 'Data final', 'Fim', 'Due Date', 'Prazo']) ||
            issue.dueDate ||
            issue.raw?.dueDate ||
            issue.raw?.attributes?.dueDate
        ) || startDate;
      const companyWork = readField(issue, ['Empresa+Obra', 'Empresa + Obra', 'Empresa Obra', 'Obra', 'Empresa']) || 'Sem obra informada';
      const location = issue.location || readField(issue, ['Localizacao', 'Localização', 'Detalhes da localizacao', 'Detalhes da localização']) || '';
      const service = readField(issue, ['Servico realizado', 'Serviço realizado', 'Atividade', 'Descricao do servico']) || issue.description || '';
      const material = readField(issue, ['Material', 'Cod. Material', 'Codigo Material', 'Cód. Material']) || '';
      const materialState = readField(issue, ['Estado do material', 'Status material']) || '';
      const materialExpanded =
        material ||
        readField(issue, [
          'Materiais',
          'Material previsto',
          'Materiais previstos',
          'Material utilizado',
          'Materiais utilizados',
          'Tipo de material'
        ]) ||
        '';
      const materialStateExpanded =
        materialState ||
        readField(issue, [
          'Situacao do material',
          'Situação do material'
        ]) ||
        '';
      const rootCause = readField(issue, ['Causa raiz', 'Causa']) || '';
      const positioning = readField(issue, ['Posicionamento', 'Localizacao GPS', 'Localização GPS']) || '';
      const principalName = issue.assignedTo || readField(issue, ['Atribuido a', 'Atribuído a', 'Responsavel', 'Responsável']);
      const followerNames = splitNames(readField(issue, ['Apoio', 'Equipe de apoio', 'Colaboradores de apoio']));
      const peopleByKey = new Map();
      const principalUser = resolveKnownProjectUser(principalName);
      if (principalUser) peopleByKey.set(principalUser.key, { ...principalUser, role: 'Equipe' });
      followerNames.forEach((name) => {
        const supportUser = resolveKnownProjectUser(name);
        if (supportUser && !peopleByKey.has(supportUser.key)) peopleByKey.set(supportUser.key, { ...supportUser, role: 'Apoio' });
      });
      const people = Array.from(peopleByKey.values());

      const common = {
        issueId: issue.id,
        issue,
        title: issue.title || 'Issue sem titulo',
        status: issue.status || 'Sem status',
        companyWork,
        location,
        service,
        material: materialExpanded,
        materialState: materialStateExpanded,
        rootCause,
        positioning,
        startDate,
        endDate
      };

      if (!people.length) {
        return [{ ...common, id: `${issue.id}:sem-responsavel`, resourceKey: 'sem-responsavel', resourceName: 'Sem responsavel', role: 'Pendente' }];
      }

      return people.map((person, index) => ({
        ...common,
        id: `${issue.id}:${person.key}:${index}`,
        resourceKey: person.key,
        resourceName: person.name,
        resourceEmail: person.email,
        role: person.role
      }));
    });
  }, [instrumentationIssues, customFieldDefinitions, resolveKnownProjectUser]);

  const instrumentDefaultPeriod = useMemo(() => {
    const datedAllocations = instrumentationAllocations.filter((allocation) => allocation.startDate);
    const today = new Date();
    const firstDate = datedAllocations.length
      ? new Date(Math.min(...datedAllocations.map((allocation) => allocation.startDate.getTime())))
      : today;
    const start = new Date(firstDate.getFullYear(), firstDate.getMonth(), 1);
    const end = new Date(firstDate.getFullYear(), firstDate.getMonth() + 3, 0);
    return { start, end };
  }, [instrumentationAllocations]);

  const instrumentPeriod = useMemo(() => {
    const customStart = parseIssueDate(instrumentStartDate);
    const customEnd = parseIssueDate(instrumentEndDate);
    if (instrumentViewMode === 'custom' && customStart && customEnd) return { start: customStart, end: customEnd };
    if (instrumentViewMode === 'week') {
      const start = customStart || instrumentDefaultPeriod.start;
      return { start, end: addDays(start, 6) };
    }
    if (instrumentViewMode === 'month') {
      const start = customStart || instrumentDefaultPeriod.start;
      return { start, end: new Date(start.getFullYear(), start.getMonth() + 1, 0) };
    }
    if (instrumentViewMode === 'quarter') {
      const start = customStart || instrumentDefaultPeriod.start;
      return { start, end: new Date(start.getFullYear(), start.getMonth() + 3, 0) };
    }
    return {
      start: customStart || instrumentDefaultPeriod.start,
      end: customEnd || instrumentDefaultPeriod.end
    };
  }, [instrumentViewMode, instrumentStartDate, instrumentEndDate, instrumentDefaultPeriod]);

  const instrumentDays = useMemo(
    () => listDatesBetween(instrumentPeriod.start, instrumentPeriod.end),
    [instrumentPeriod]
  );

  useEffect(() => {
    if (selectedBusinessUnit !== 'G5 Instrumentos' || !selectedProjectId) {
      setInstrumentResourceOverrides({});
      return;
    }

    const cacheKey = `central-g5-instrumentos-recursos:${selectedProjectId}`;
    try {
      setInstrumentResourceOverrides(JSON.parse(localStorage.getItem(cacheKey) || '{}'));
    } catch (storageError) {
      setInstrumentResourceOverrides({});
    }
  }, [selectedBusinessUnit, selectedProjectId]);

  const instrumentResources = useMemo(() => {
    const resources = new Map();
    projectUsers.forEach((projectUser) => {
      const key = normalizeText(projectUser.name || projectUser.email || projectUser.id);
      if (!key) return;
      resources.set(key, {
        key,
        name: projectUser.name || projectUser.email || projectUser.id,
        email: projectUser.email || '',
        company: projectUser.company || projectUser.companyName || '',
        operationalRole: projectUser.role || projectUser.projectRole || '',
        base: '',
        resourceType: 'Equipe',
        active: true,
        capacityDaily: 1
      });
    });
    instrumentationAllocations.forEach((allocation) => {
      const rawName = String(allocation.resourceName || '').trim();
      const looksLikeAutodeskId = /^[a-z0-9]{10,}$/i.test(rawName) && !rawName.includes(' ');
      if (!resources.has(allocation.resourceKey) && (allocation.resourceKey === 'sem-responsavel' || !looksLikeAutodeskId)) {
        resources.set(allocation.resourceKey, {
          key: allocation.resourceKey,
          name: allocation.resourceName,
          email: '',
          company: '',
          operationalRole: '',
          base: '',
          resourceType: allocation.role === 'Apoio' ? 'Apoio' : 'Campo',
          active: true,
          capacityDaily: 1
        });
      }
    });

    return Array.from(resources.values())
      .map((resource) => ({ ...resource, ...(instrumentResourceOverrides[resource.key] || {}) }))
      .filter((resource) => resource.active !== false)
      .sort((firstResource, secondResource) => String(firstResource.name).localeCompare(String(secondResource.name), 'pt-BR', { sensitivity: 'base' }));
  }, [projectUsers, instrumentationAllocations, instrumentResourceOverrides]);

  const filteredInstrumentationAllocations = useMemo(() => {
    const periodStart = instrumentPeriod.start?.getTime();
    const periodEnd = instrumentPeriod.end?.getTime();
    return instrumentationAllocations.filter((allocation) => {
      if (allocation.endDate && allocation.endDate.getTime() < periodStart) return false;
      if (allocation.startDate && allocation.startDate.getTime() > periodEnd) return false;
      if (instrumentWorkFilter !== 'all' && allocation.companyWork !== instrumentWorkFilter) return false;
      if (instrumentResourceFilter !== 'all' && allocation.resourceKey !== instrumentResourceFilter) return false;
      if (instrumentStatusFilter !== 'all' && allocation.status !== instrumentStatusFilter) return false;
      if (instrumentLocationFilter !== 'all' && (allocation.location || 'Sem localizacao') !== instrumentLocationFilter) return false;
      if (instrumentMaterialFilter !== 'all' && (allocation.material || 'Sem material informado') !== instrumentMaterialFilter) return false;
      const resource = instrumentResources.find((item) => item.key === allocation.resourceKey);
      if (instrumentFunctionFilter !== 'all' && (resource?.operationalRole || 'Sem funcao definida') !== instrumentFunctionFilter) return false;
      if (instrumentResourceTypeFilter !== 'all' && (resource?.resourceType || 'Equipe') !== instrumentResourceTypeFilter) return false;
      return true;
    });
  }, [
    instrumentationAllocations,
    instrumentPeriod,
    instrumentWorkFilter,
    instrumentResourceFilter,
    instrumentStatusFilter,
    instrumentLocationFilter,
    instrumentMaterialFilter,
    instrumentResources,
    instrumentFunctionFilter,
    instrumentResourceTypeFilter
  ]);

  const instrumentAllocationByDay = useMemo(() => {
    const byKey = new Map();
    filteredInstrumentationAllocations.forEach((allocation) => {
      if (!allocation.startDate || !allocation.endDate) return;
      instrumentDays.forEach((day) => {
        if (day < allocation.startDate || day > allocation.endDate) return;
        const key = `${allocation.resourceKey}:${formatInputDate(day)}`;
        const current = byKey.get(key) || [];
        current.push(allocation);
        byKey.set(key, current);
      });
    });
    return byKey;
  }, [filteredInstrumentationAllocations, instrumentDays]);

  const selectedInstrumentAllocation = useMemo(
    () => filteredInstrumentationAllocations.find((allocation) => allocation.id === selectedInstrumentAllocationId) || null,
    [filteredInstrumentationAllocations, selectedInstrumentAllocationId]
  );

  const instrumentMonthBoards = useMemo(() => {
    const months = new Map();
    instrumentDays.forEach((day) => {
      const key = `${day.getFullYear()}-${String(day.getMonth() + 1).padStart(2, '0')}`;
      if (!months.has(key)) {
        months.set(key, {
          key,
          label: day.toLocaleDateString('pt-BR', { month: 'long', year: 'numeric' }),
          firstDay: new Date(day.getFullYear(), day.getMonth(), 1),
          days: []
        });
      }
      months.get(key).days.push(day);
    });

    return Array.from(months.values()).slice(0, 3).map((month) => ({
      ...month,
      leadingBlanks: month.firstDay.getDay(),
      allocations: month.days.reduce((accumulator, day) => {
        const dayKey = formatInputDate(day);
        accumulator[dayKey] = filteredInstrumentationAllocations.filter((allocation) => {
          if (!allocation.startDate || !allocation.endDate) return false;
          return day >= allocation.startDate && day <= allocation.endDate;
        });
        return accumulator;
      }, {})
    }));
  }, [instrumentDays, filteredInstrumentationAllocations]);

  const instrumentFilterOptions = useMemo(() => {
    const unique = (values) => Array.from(new Set(values.filter(Boolean))).sort((first, second) => String(first).localeCompare(String(second), 'pt-BR', { sensitivity: 'base' }));
    return {
      works: unique(instrumentationAllocations.map((allocation) => allocation.companyWork)),
      resources: instrumentResources.map((resource) => ({ value: resource.key, label: resource.name })),
      functions: unique(instrumentResources.map((resource) => resource.operationalRole || 'Sem funcao definida')),
      statuses: unique(instrumentationAllocations.map((allocation) => allocation.status)),
      locations: unique(instrumentationAllocations.map((allocation) => allocation.location || 'Sem localizacao')),
      resourceTypes: unique(instrumentResources.map((resource) => resource.resourceType || 'Equipe')),
      materials: unique(instrumentationAllocations.map((allocation) => allocation.material || 'Sem material informado'))
    };
  }, [instrumentationAllocations, instrumentResources]);

  const instrumentationIndicators = useMemo(() => {
    const conflictKeys = Array.from(instrumentAllocationByDay.values()).filter((items) => items.length > 1);
    const uniqueWorks = new Set(filteredInstrumentationAllocations.map((allocation) => allocation.companyWork).filter(Boolean));
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    return {
      activeResources: instrumentResources.length,
      works: uniqueWorks.size,
      allocations: filteredInstrumentationAllocations.length,
      materials: new Set(filteredInstrumentationAllocations.map((allocation) => allocation.material).filter(Boolean)).size,
      conflicts: conflictKeys.length,
      missingStart: instrumentationIssues.filter((issue) => !parseIssueDate(getCustomFieldValueByAliases(issue, customFieldDefinitions, ['Data de inicio', 'Data de início']) || issue.startDate)).length,
      missingEnd: instrumentationIssues.filter((issue) => !parseIssueDate(getCustomFieldValueByAliases(issue, customFieldDefinitions, ['Data de vencimento', 'Due Date', 'Prazo']) || issue.dueDate)).length,
      missingResponsible: instrumentationIssues.filter((issue) => !issue.assignedTo).length,
      overdue: instrumentationIssues.filter((issue) => {
        const dueDate = parseIssueDate(issue.dueDate || getCustomFieldValueByAliases(issue, customFieldDefinitions, ['Data de vencimento', 'Prazo']));
        return dueDate && dueDate < today && !['closed', 'completed', 'concluido', 'concluida'].includes(normalizeText(issue.status));
      }).length
    };
  }, [instrumentAllocationByDay, filteredInstrumentationAllocations, instrumentResources, instrumentationIssues, customFieldDefinitions]);

  const instrumentWorks = useMemo(() => {
    const works = new Map();
    instrumentationAllocations.forEach((allocation) => {
      const key = allocation.companyWork || 'Sem obra informada';
      const current = works.get(key) || {
        key,
        name: key,
        issueIds: new Set(),
        resources: new Set(),
        support: new Set(),
        materials: new Set(),
        overdue: 0,
        missingEnd: 0,
        startDates: [],
        endDates: []
      };
      current.issueIds.add(allocation.issueId);
      current.resources.add(allocation.resourceName);
      if (allocation.role === 'Apoio') current.support.add(allocation.resourceName);
      if (allocation.material) current.materials.add(allocation.material);
      if (allocation.startDate) current.startDates.push(allocation.startDate);
      if (allocation.endDate) current.endDates.push(allocation.endDate);
      if (!allocation.endDate) current.missingEnd += 1;
      works.set(key, current);
    });

    return Array.from(works.values()).map((work) => {
      const start = work.startDates.length ? new Date(Math.min(...work.startDates.map((date) => date.getTime()))) : null;
      const end = work.endDates.length ? new Date(Math.max(...work.endDates.map((date) => date.getTime()))) : null;
      return {
        ...work,
        issues: work.issueIds.size,
        team: work.resources.size,
        supportCount: work.support.size,
        materialCount: work.materials.size,
        start,
        end
      };
    }).sort((first, second) => String(first.name).localeCompare(String(second.name), 'pt-BR', { sensitivity: 'base' }));
  }, [instrumentationAllocations]);

  const instrumentEquipmentItems = useMemo(() => {
    const readField = (issue, aliases) => getCustomFieldValueByAliases(issue, customFieldDefinitions, aliases) || getIssueCustomValue(issue, aliases) || '';
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const soonLimit = addDays(today, 30);

    return instrumentationIssues
      .map((issue) => {
        const code = readField(issue, ['Codigo do equipamento', 'Codigo equipamento', 'Cod. equipamento', 'Patrimonio', 'Patrimônio']);
        const name = readField(issue, ['Equipamento', 'Nome do equipamento', 'Instrumento', 'Recurso equipamento']);
        const status = readField(issue, ['Status do equipamento', 'Status equipamento', 'Situacao do equipamento', 'Situação do equipamento']) || 'Nao informado';
        const obra = readField(issue, ['Empresa+Obra', 'Empresa + Obra', 'Obra', 'Empresa']) || 'Sem obra informada';
        const responsible = readField(issue, ['Responsavel equipamento', 'Responsável equipamento', 'Responsavel pela retirada', 'Responsável pela retirada']) || issue.assignedTo || '';
        const returnDate = parseIssueDate(readField(issue, ['Data prevista de retorno', 'Previsao de retorno', 'Previsão de retorno']));
        const calibrationDate = parseIssueDate(readField(issue, ['Data da proxima calibracao', 'Data da próxima calibração', 'Proxima calibracao', 'Próxima calibração']));
        if (!code && !name && normalizeText(status) === 'nao informado') return null;
        return {
          issue,
          code: code || issue.displayId || issue.id,
          name: name || issue.title || 'Equipamento informado no issue',
          status,
          obra,
          responsible: responsible || 'Sem responsavel',
          returnDate,
          calibrationDate,
          returnOverdue: returnDate && returnDate < today,
          calibrationOverdue: calibrationDate && calibrationDate < today,
          calibrationSoon: calibrationDate && calibrationDate >= today && calibrationDate <= soonLimit
        };
      })
      .filter(Boolean);
  }, [instrumentationIssues, customFieldDefinitions]);

  const instrumentMaterialItems = useMemo(() => {
    const readField = (issue, aliases) => getCustomFieldValueByAliases(issue, customFieldDefinitions, aliases) || getIssueCustomValue(issue, aliases) || '';
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    return instrumentationIssues
      .map((issue) => {
        const material =
          readField(issue, ['Material', 'Materiais', 'Material previsto', 'Material utilizado', 'Tipo de material']) || '';
        const code = readField(issue, ['Cod. Material', 'Codigo Material', 'Codigo do item', 'Código do item']);
        const unit = readField(issue, ['Unidade', 'Unidade padrao', 'Unidade padrão']);
        const quantity = readField(issue, ['Quantidade', 'Quantidade prevista', 'Qtd', 'QTD']);
        const compraStatus = readField(issue, ['Status de compra', 'Status compra', 'Compra', 'Situacao compra']) || 'Nao solicitado';
        const deliveryDate = parseIssueDate(readField(issue, ['Data prevista de entrega', 'Previsao de entrega', 'Previsão de entrega']));
        const obra = readField(issue, ['Empresa+Obra', 'Empresa + Obra', 'Obra', 'Empresa']) || 'Sem obra informada';
        if (!material && !code) return null;
        return {
          issue,
          code: code || '-',
          material: material || 'Material sem descricao',
          obra,
          unit: unit || 'Sem unidade',
          quantity: quantity || 'Sem quantidade',
          compraStatus,
          deliveryDate,
          late: deliveryDate && deliveryDate < today && !['recebido', 'comprado', 'cancelado'].includes(normalizeText(compraStatus)),
          classified: Boolean(code || readField(issue, ['Codigo ABNT NBR 15965', 'Código ABNT NBR 15965', 'Codigo NBR 15965']))
        };
      })
      .filter(Boolean);
  }, [instrumentationIssues, customFieldDefinitions]);

  const instrumentAlerts = useMemo(() => {
    const alerts = [];
    if (instrumentationIndicators.conflicts > 0) alerts.push({ level: 'Critica', area: 'Equipe', text: `${instrumentationIndicators.conflicts} dias com conflito de alocacao.` });
    if (instrumentationIndicators.missingResponsible > 0) alerts.push({ level: 'Alta', area: 'Equipe', text: `${instrumentationIndicators.missingResponsible} issues sem responsavel principal.` });
    if (instrumentationIndicators.missingStart > 0) alerts.push({ level: 'Media', area: 'Equipe', text: `${instrumentationIndicators.missingStart} issues sem data de inicio.` });
    if (instrumentationIndicators.missingEnd > 0) alerts.push({ level: 'Media', area: 'Equipe', text: `${instrumentationIndicators.missingEnd} issues sem data de vencimento.` });
    instrumentEquipmentItems.filter((item) => item.returnOverdue).slice(0, 5).forEach((item) => alerts.push({ level: 'Alta', area: 'Equipamento', text: `${item.name} com retorno vencido em ${formatDate(item.returnDate)}.` }));
    instrumentEquipmentItems.filter((item) => item.calibrationOverdue).slice(0, 5).forEach((item) => alerts.push({ level: 'Critica', area: 'Equipamento', text: `${item.name} com calibracao vencida.` }));
    instrumentMaterialItems.filter((item) => item.late).slice(0, 5).forEach((item) => alerts.push({ level: 'Alta', area: 'Material', text: `${item.material} com entrega atrasada para ${item.obra}.` }));
    instrumentMaterialItems.filter((item) => !item.classified).slice(0, 5).forEach((item) => alerts.push({ level: 'Baixa', area: 'Material', text: `${item.material} ainda sem classificacao padronizada.` }));
    return alerts;
  }, [instrumentationIndicators, instrumentEquipmentItems, instrumentMaterialItems]);

  const instrumentPurchaseStats = useMemo(() => {
    const normalizePurchase = (value) => normalizeText(value || 'Nao solicitado');
    return {
      pending: instrumentMaterialItems.filter((item) => ['nao solicitado', 'em validacao tecnica', 'aguardando cotacao', 'em cotacao', 'aguardando aprovacao'].includes(normalizePurchase(item.compraStatus))).length,
      quoted: instrumentMaterialItems.filter((item) => normalizePurchase(item.compraStatus).includes('cotacao')).length,
      bought: instrumentMaterialItems.filter((item) => ['pedido emitido', 'comprado', 'parcialmente recebido'].includes(normalizePurchase(item.compraStatus))).length,
      received: instrumentMaterialItems.filter((item) => normalizePurchase(item.compraStatus) === 'recebido').length
    };
  }, [instrumentMaterialItems]);

  const eapResultSummary = useMemo(() => {
    if (!eapResults.length) return null;
    const counters = eapResults.reduce(
      (accumulator, row) => {
        const status = (row.status || '').toLowerCase();
        if (status === 'criado') accumulator.created += 1;
        else if (status === 'duplicado') accumulator.duplicated += 1;
        else if (status === 'erro') accumulator.errors += 1;
        else if (status === 'simulado') accumulator.simulated += 1;
        else accumulator.review += 1;
        return accumulator;
      },
      { created: 0, duplicated: 0, errors: 0, simulated: 0, review: 0 }
    );
    return { total: eapResults.length, ...counters };
  }, [eapResults]);

  const applicableCustomFieldDefinitions = useMemo(() => {
    return customFieldDefinitions.filter((field) => {
      if (!field.appliesTo?.length) return false;

      return field.appliesTo.some((mapping) => {
        const mappingType = normalizeText(mapping.mappedItemType);
        const isSelectedType = mapping.mappedItemId === newInterfaceIssue.issueTypeId && mappingType.includes('type');
        const isSelectedSubtype =
          newInterfaceIssue.issueSubtypeId &&
          mapping.mappedItemId === newInterfaceIssue.issueSubtypeId &&
          mappingType.includes('subtype');

        return isSelectedType || isSelectedSubtype;
      });
    });
  }, [customFieldDefinitions, newInterfaceIssue.issueTypeId, newInterfaceIssue.issueSubtypeId]);

  const visibleIssues = useMemo(() => {
    let filteredIssues = issues;

    if (issueFilter === 'overdue') {
      filteredIssues = issues.filter(isOverdue);
    }

    if (categoryFilter !== 'all') {
      filteredIssues = filteredIssues.filter((issue) => (issue.category || 'Sem categoria') === categoryFilter);
    }

    if (typeFilter !== 'all') {
      filteredIssues = filteredIssues.filter((issue) => (issue.issueType || 'Sem tipo') === typeFilter);
    }

    return [...filteredIssues].sort(sortByName);
  }, [issues, issueFilter, categoryFilter, typeFilter]);

  const filterOptions = useMemo(() => {
    const buildOptions = (fieldName, fallbackLabel) =>
      [...new Set(issues.map((issue) => issue[fieldName] || fallbackLabel))]
        .filter(Boolean)
        .sort((firstValue, secondValue) => String(firstValue).localeCompare(String(secondValue), 'pt-BR', { sensitivity: 'base' }));

    return {
      categories: buildOptions('category', 'Sem categoria'),
      types: buildOptions('issueType', 'Sem tipo')
    };
  }, [issues]);

  const metrics = useMemo(() => {
    return {
      total: issues.length,
      visible: visibleIssues.length,
      open: issues.filter(isOpenIssue).length,
      overdue: issues.filter(isOverdue).length
    };
  }, [issues, visibleIssues]);

  function getEapParts(value) {
    return String(value || '')
      .trim()
      .split('.')
      .map((part) => part.trim())
      .filter(Boolean)
      .map((part) => {
        const numeric = Number(part.replace(/[^0-9-]/g, ''));
        return Number.isFinite(numeric) ? numeric : part;
      });
  }

  function compareEapCodes(firstValue, secondValue) {
    const firstParts = getEapParts(firstValue);
    const secondParts = getEapParts(secondValue);
    const maxLength = Math.max(firstParts.length, secondParts.length);
    for (let index = 0; index < maxLength; index += 1) {
      const firstPart = firstParts[index];
      const secondPart = secondParts[index];
      if (firstPart === undefined) return -1;
      if (secondPart === undefined) return 1;
      if (typeof firstPart === 'number' && typeof secondPart === 'number' && firstPart !== secondPart) return firstPart - secondPart;
      const textCompare = String(firstPart).localeCompare(String(secondPart), 'pt-BR', { numeric: true, sensitivity: 'base' });
      if (textCompare) return textCompare;
    }
    return 0;
  }

  function renderEapLinkedIssueTitle(row) {
    const level = Math.max(0, getEapParts(row?.eapVinculada).length - 1);
    const title = row?.titulo || row?.issue?.title || 'Issue sem título';
    return (
      <span className={`eap-linked-title eap-depth-${Math.min(level, 5)}`} title={title}>
        {level > 0 && <span className="eap-branch" aria-hidden="true">↳</span>}
        <span className={`eap-node-type ${row?.isMarcoContratual ? 'is-marco' : 'is-task'}`} aria-hidden="true">{row?.isMarcoContratual ? 'M' : 'T'}</span>
        <span>{title}</span>
      </span>
    );
  }

  const cronogramaRows = useMemo(() => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const lookup = buildIssueLookupIndexes(issues);
    const finalStatuses = ['concluido', 'concluído', 'fechado', 'finalizado', 'aprovado', 'done', 'closed', 'completed'];
    const isFinalStatus = (value) => finalStatuses.some((status) => normalizeText(value).includes(status));
    const calculateStatusProgress = (value) => {
      const normalized = normalizeText(value);
      if (!normalized) return 0;
      if (isFinalStatus(value)) return 100;
      if (normalized.includes('revis') || normalized.includes('analise') || normalized.includes('análise') || normalized.includes('valid')) return 75;
      if (normalized.includes('andamento') || normalized.includes('progress') || normalized.includes('execu') || normalized.includes('em curso')) return 50;
      return 0;
    };
    console.info('[Modulo 8][Cronograma] Total de issues carregados:', issues.length);
    if (issues[0]) console.info('[Modulo 8][Cronograma] Exemplo de issue retornado:', issues[0]);
    const predecessorFieldDefinition = (customFieldDefinitions || []).find((field) =>
      [field.name, field.title, field.displayName, field.label, field.id, field.definitionId]
        .filter(Boolean)
        .map(normalizeFieldKey)
        .includes(normalizeFieldKey('Predecessor'))
    );
    console.info('[Modulo 8][Cronograma] Campos personalizados encontrados:', (customFieldDefinitions || []).map((field) => field.name || field.title || field.id));
    console.info('[Modulo 8][Cronograma] Campo Predecessor (fieldId/definitionId):', predecessorFieldDefinition?.id || predecessorFieldDefinition?.definitionId || 'nao encontrado');
    const rows = issues.map((issue) => {
      const eapRaw = getCustomFieldValueByAliases(issue, customFieldDefinitions, ['Código EAP', 'Codigo EAP', 'EAP Vinculada', 'EAP']);
      const codigoMarco = getCustomFieldValueByAliases(issue, customFieldDefinitions, CODIGO_MARCO_FIELD_ALIASES) || '';
      const marcoContratual = getCustomFieldValueByAliases(issue, customFieldDefinitions, [
        'Marco na planilha',
        'Marco da planilha',
        'Marco Contratual',
        'Marco'
      ]) || codigoMarco || 'Sem marco';
      const eapVinculada = eapRaw || (marcoContratual !== 'Sem marco' ? marcoContratual : (issue.displayId || issue.id || 'Sem EAP'));
      const eapNivel1 = String(eapVinculada || '').trim().split('.').filter(Boolean)[0] || '';
      const titulo = issue.title || 'Issue sem titulo';
      const dataContratual = getCustomDateFieldValue(issue, ['Data Contratual', 'Data contratual']) || '';
      const dataPrevistaG5 = getCustomDateFieldValue(issue, ['Data Prevista G5', 'Data prevista G5']) || '';
      const inicioPrevisto = getCustomDateFieldValue(issue, scheduleFieldAliases.inicioPlanejado) || issue.startDate || issue.raw?.startDate || '';
      const terminoPrevisto = getCustomDateFieldValue(issue, scheduleFieldAliases.terminoPlanejado) || issue.dueDate || issue.deadline || issue.raw?.dueDate || '';
      const inicioReal = getCustomDateFieldValue(issue, ['Inicio Real', 'Início Real', 'Data Inicio Real', 'Data Início Real']) || '';
      const terminoReal = getCustomDateFieldValue(issue, ['Termino Real', 'Término Real', 'Data Termino Real', 'Data Término Real']) || '';
      const startDate = issue.startDate || getCustomFieldValue(issue, customFieldDefinitions, 'Start Date');
      const assignedRaw = issue.assignedTo || issue.assignee || issue.assignedUsers || issue.raw?.assignedTo;
      const atribuidoA = Array.isArray(assignedRaw) ? assignedRaw.map(toDisplayCustomFieldValue).filter(Boolean).join(', ') : toDisplayCustomFieldValue(assignedRaw) || 'Sem atribuição';
      const predecessorFieldValue = getCustomFieldValueByAliases(issue, customFieldDefinitions, ['Predecessor']);
      const tempoAtividade = getCustomFieldValueByAliases(issue, customFieldDefinitions, [
        'Dias previstos para atividade',
        'Dias Previsto para Atividade',
        'Dias Previstos para Atividade',
        'Dias previstos atividade',
        'Dias previstos',
        'Tempo para Atividade',
        'Tempo para atividade'
      ]);
      const terminoPrevistoCalculado = computeEndDateFromStartAndDays(cronogramaEdits[issue.id]?.inicioPrevisto ?? inicioPrevisto, cronogramaEdits[issue.id]?.tempoAtividade ?? tempoAtividade);
      const terminoPrevistoFinal = terminoPrevistoCalculado || terminoPrevisto;
      const predecessorInputValue = normalizePredecessorInput(cronogramaEdits[issue.id]?.predecessor ?? predecessorFieldValue);
      const predecessorIds = parsePredecessorIds(predecessorInputValue);
      const resolvedFromField = predecessorIds
        .map((predId) => lookup.byDisplayId.get(normalizeText(predId)) || lookup.byId.get(String(predId)) || lookup.byId.get(normalizeText(predId)))
        .filter(Boolean);
      const fallbackReferences = resolveIssueReferences(issue, lookup);
      const predecessors = predecessorIds.length ? resolvedFromField : fallbackReferences;
      if (predecessorInputValue) {
        console.info('[Modulo 8][Cronograma] Predecessor por issue', { issueId: issue.id, predecessorOriginal: predecessorFieldValue, predecessorEditado: predecessorInputValue, resolved: resolvedFromField.map((item) => item.displayId || item.id) });
        predecessorIds.forEach((predId) => {
          if (!(lookup.byDisplayId.get(normalizeText(predId)) || lookup.byId.get(String(predId)) || lookup.byId.get(normalizeText(predId)))) console.warn(`[Modulo 8][Cronograma] ID de predecessor não encontrado: ${predId} (issue ${issue.displayId || issue.id})`);
        });
      }
      const statusEntrega = getCustomFieldValueByAliases(issue, customFieldDefinitions, scheduleFieldAliases.statusEntrega) || '';
      const statusValue = cronogramaEdits[issue.id]?.status ?? statusEntrega ?? issue.status;
      const concluded = isFinalStatus(statusValue) || isScheduleCompleted(issue, customFieldDefinitions, { statusEntrega });
      const start = parseIssueDate(startDate);
      const prevista = parseIssueDate(terminoPrevistoFinal) || parseIssueDate(dataPrevistaG5);
      const contratual = parseIssueDate(dataContratual);
      const inicioRealValue = cronogramaEdits[issue.id]?.inicioReal ?? inicioReal;
      const terminoRealValue = cronogramaEdits[issue.id]?.terminoReal ?? terminoReal;
      const terminoRealDate = parseIssueDate(terminoRealValue);
      const dataBaseAtraso = prevista || contratual;
      const overdue = !!(dataBaseAtraso && dataBaseAtraso < today && !concluded);
      const blocked = predecessors.some((predecessor) => !isFinalStatus(predecessor.status));
      const semData = !prevista && !contratual;
      const durationDays = prevista ? Math.floor((prevista - today) / 86400000) : null;
      const delayDays = terminoRealDate && prevista
        ? calculateBusinessDiffDays(terminoPrevistoFinal, terminoRealValue)
        : (!concluded && dataBaseAtraso && dataBaseAtraso < today
          ? Math.max(1, calculateBusinessDiffDays(dataBaseAtraso, today) || calculateDateDiffDays(dataBaseAtraso, today) || 1)
          : null);
      const tipoValue = [issue.issueType, issue.type, issue.issueSubtype, issue.category, getIssueCustomValue(issue, ['Tipo', 'Categoria', 'Issue Type'])].filter(Boolean).join(' ');
      const isMarcoContratual = normalizeText(tipoValue).includes('marco contratual/entrega')
        || (normalizeText(tipoValue).includes('marco contratual') && normalizeText(tipoValue).includes('entrega'))
        || normalizeText(tipoValue).includes('marco contratual')
        || (marcoContratual !== 'Sem marco' && !eapRaw)
        || /^mc\d+/i.test(titulo.trim());
      const statusConsolidado = blocked
        ? 'Bloqueado'
        : overdue
          ? 'Atrasado'
          : concluded
            ? 'Concluído'
            : semData
              ? 'Sem status'
              : 'Em andamento';

      return {
        issue, eapVinculada, codigoMarco, marcoContratual, titulo, dataContratual, atribuidoA, dataPrevistaG5, inicioPrevisto, terminoPrevisto: terminoPrevistoFinal, terminoPrevistoOriginal: terminoPrevisto, inicioReal: inicioRealValue, terminoReal: terminoRealValue, tempoAtividade, startDate, statusEntrega,
        predecessors, predecessorFieldValue: predecessorInputValue, predecessorIds, concluded, durationDays, delayDays, overdue, blocked, semData, statusConsolidado, isMarcoContratual, eapNivel1, statusValue
      };
    });
    const calculateMarcoExecution = (marco) => {
      const filhos = rows.filter((row) => !row.isMarcoContratual && normalizeText(row.marcoContratual) === normalizeText(marco));
      if (!filhos.length) {
        const marcoRow = rows.find((row) => row.isMarcoContratual && normalizeText(row.marcoContratual) === normalizeText(marco));
        return calculateRealExecution(marcoRow);
      }
      const progressoTotal = filhos.reduce((sum, row) => sum + calculateRealExecution(row), 0);
      return Math.round(progressoTotal / filhos.length);
    };
    function calculateRealExecution(row) {
      if (!row) return 0;
      const statusProgress = getScheduleProgressFromDeliveryStatus(row.statusEntrega || row.statusValue || row.issue?.status);
      if (statusProgress > 0) return statusProgress;
      if (isFinalStatus(row.statusValue) || row.concluded) return 100;
      const realStart = parseIssueDate(row.inicioReal);
      if (!realStart) return 0;
      const realEnd = parseIssueDate(row.terminoReal) || today;
      const plannedStart = parseIssueDate(row.inicioPrevisto) || parseIssueDate(row.dataContratual) || realStart;
      const plannedEnd = parseIssueDate(row.terminoPrevisto) || plannedStart;
      const durationFromField = parseDurationDays(row.tempoAtividade);
      const plannedDays = Math.max(
        1,
        durationFromField || calculateBusinessDiffDays(plannedStart, plannedEnd) || calculateDateDiffDays(plannedStart, plannedEnd) || 1
      );
      const realDays = Math.max(0, calculateBusinessDiffDays(realStart, realEnd) || calculateDateDiffDays(realStart, realEnd) || 0);
      return Math.max(1, Math.min(99, Math.round((realDays / plannedDays) * 100)));
    }
    rows.forEach((row) => {
      const execution = row.isMarcoContratual ? calculateMarcoExecution(row.marcoContratual) : calculateRealExecution(row);
      row.executado = execution;
      row.predecessoresTexto = row.predecessorIds.length ? row.predecessorIds.map((predId) => {
        const predecessor = row.predecessors.find((item) => String(item.displayId || item.id) === String(predId) || String(item.id) === String(predId));
        if (!predecessor) return `ID não encontrado: ${predId}`;
        const predEap = getCustomFieldValueByAliases(predecessor, customFieldDefinitions, ['Código EAP', 'Codigo EAP', 'EAP Vinculada', 'EAP']);
        return `${predId} — ${predEap || predecessor.title || predecessor.displayId || predecessor.id}`;
      }).join(', ') : '-';
      row.predecessoresTexto = row.predecessors.map((predecessor) => {
        const predEap = getCustomFieldValueByAliases(predecessor, customFieldDefinitions, ['Código EAP', 'Codigo EAP', 'EAP Vinculada', 'EAP']);
        return predEap || predecessor.title || predecessor.displayId || '-';
      }).join(', ') || '-';
    });
    console.info('[Modulo 8][Cronograma] Definições de campos personalizados:', customFieldDefinitions);
    console.info('[Modulo 8][Cronograma] Lista de campos personalizados encontrados:', rows.slice(0, 5).map((row) => extractIssueCustomAttributeEntries(row.issue).map((entry) => entry.names[0])));
    console.info('[Modulo 8][Cronograma] Resultado de mapeamentos', rows.slice(0, 5).map((row) => ({
      issue: row.issue.displayId || row.issue.id,
      eap: row.eapVinculada,
      marco: row.marcoContratual,
      dataContratual: row.dataContratual,
      dataPrevistaG5: row.dataPrevistaG5,
      inicioPrevisto: row.inicioPrevisto,
      terminoPrevisto: row.terminoPrevisto,
      inicioReal: row.inicioReal,
      terminoReal: row.terminoReal,
      predecessors: row.predecessors.length
    })));
    return rows.sort((a, b) =>
      compareEapCodes(a.eapVinculada, b.eapVinculada)
      || String(a.marcoContratual).localeCompare(String(b.marcoContratual), 'pt-BR', { numeric: true, sensitivity: 'base' })
      || (a.isMarcoContratual === b.isMarcoContratual ? 0 : (a.isMarcoContratual ? -1 : 1))
      || String(a.dataContratual).localeCompare(String(b.dataContratual))
      || String(a.dataPrevistaG5).localeCompare(String(b.dataPrevistaG5))
      || String(a.titulo).localeCompare(String(b.titulo), 'pt-BR', { sensitivity: 'base' })
    );
  }, [issues, customFieldDefinitions, cronogramaEdits]);

  const cronogramaPendingChangesCount = useMemo(() => Object.keys(cronogramaEdits).length, [cronogramaEdits]);
  const cronogramaPageSize = 10;
  const cronogramaRowsFiltered = useMemo(() => {
    if (!cronogramaDashboardFilter) return cronogramaRows;
    return cronogramaRows.filter((row) => {
      if (cronogramaDashboardFilter.kind === 'status') return cronogramaDashboardFilter.value === (row.statusValue || 'Sem status');
      if (cronogramaDashboardFilter.kind === 'tipo') return cronogramaDashboardFilter.value === (row.issue.issueType || row.issue.type || row.issue.category || 'Tipo não identificado');
      if (cronogramaDashboardFilter.kind === 'marco') return normalizeText(cronogramaDashboardFilter.value) === normalizeText(row.marcoContratual || 'Sem marco');
      if (cronogramaDashboardFilter.kind === 'prazo') {
        if (cronogramaDashboardFilter.value === 'Concluídos') return row.concluded;
        if (cronogramaDashboardFilter.value === 'Atrasados') return row.overdue && !row.concluded;
        if (cronogramaDashboardFilter.value === 'Sem data') return !row.dataPrevistaG5 && !row.dataContratual && !row.concluded;
        if (cronogramaDashboardFilter.value === 'Em dia') return !row.concluded && !row.overdue && (row.dataPrevistaG5 || row.dataContratual);
      }
      return true;
    });
  }, [cronogramaRows, cronogramaDashboardFilter]);
  const cronogramaRowsSorted = useMemo(() => {
    const getSortValue = (row, key) => {
      if (key === 'status') return row.statusValue || row.issue.status || '';
      if (key === 'executado' || key === 'delayDays') return Number(row[key] || 0);
      const value = row[key];
      const parsedDate = parseIssueDate(value);
      if (parsedDate) return parsedDate.getTime();
      return String(value || '').toLocaleLowerCase('pt-BR');
    };
    const direction = cronogramaSort.direction === 'desc' ? -1 : 1;
    return [...cronogramaRowsFiltered].sort((firstRow, secondRow) => {
      if (cronogramaSort.key === 'eapVinculada') {
        return compareEapCodes(firstRow.eapVinculada, secondRow.eapVinculada) * direction;
      }
      const firstValue = getSortValue(firstRow, cronogramaSort.key);
      const secondValue = getSortValue(secondRow, cronogramaSort.key);
      if (typeof firstValue === 'number' && typeof secondValue === 'number') {
        return (firstValue - secondValue) * direction;
      }
      return String(firstValue).localeCompare(String(secondValue), 'pt-BR', { numeric: true, sensitivity: 'base' }) * direction;
    });
  }, [cronogramaRowsFiltered, cronogramaSort]);
  const cronogramaTotalPages = useMemo(() => Math.max(1, Math.ceil(cronogramaRowsFiltered.length / cronogramaPageSize)), [cronogramaRowsFiltered.length]);
  const pagedCronogramaRows = useMemo(() => {
    const start = (cronogramaCurrentPage - 1) * cronogramaPageSize;
    return cronogramaRowsFiltered.slice(start, start + cronogramaPageSize);
  }, [cronogramaRowsFiltered, cronogramaCurrentPage]);
  const cronogramaDashboard = useMemo(() => {
    const baseRows = cronogramaRows;
    const activityRows = baseRows.filter((row) => !row.isMarcoContratual);
    const byKey = (items, fn) => {
      const map = new Map();
      items.forEach((item) => map.set(fn(item), (map.get(fn(item)) || 0) + 1));
      return Array.from(map.entries()).map(([name, value]) => ({ name, value })).sort((a, b) => b.value - a.value);
    };
    const statusData = byKey(baseRows, (row) => row.statusValue || 'Sem status');
    const marcoData = byKey(baseRows, (row) => row.marcoContratual || 'Sem marco').sort((a, b) => String(a.name).localeCompare(String(b.name), 'pt-BR', { numeric: true, sensitivity: 'base' }));
    const tipoData = byKey(baseRows, (row) => row.issue.issueType || row.issue.type || row.issue.category || 'Tipo não identificado');
    const prazoData = [
      { name: 'Concluídos', value: baseRows.filter((row) => row.concluded).length, color: '#0f7c90' },
      { name: 'Atrasados', value: baseRows.filter((row) => row.overdue && !row.concluded).length, color: '#d97373' },
      { name: 'Sem data', value: baseRows.filter((row) => !row.dataPrevistaG5 && !row.dataContratual && !row.concluded).length, color: '#d2d6de' },
      { name: 'Em dia', value: baseRows.filter((row) => !row.concluded && !row.overdue && (row.dataPrevistaG5 || row.dataContratual)).length, color: '#2a6aa5' }
    ];
    const execucaoPorMarco = marcoData.map((marco) => {
      const rows = activityRows.filter((row) => normalizeText(row.marcoContratual || 'Sem marco') === normalizeText(marco.name));
      return { name: marco.name, concluidos: rows.filter((r) => r.concluded).length, atrasados: rows.filter((r) => r.overdue && !r.concluded).length, pendentes: rows.filter((r) => !r.concluded && !r.overdue).length, total: rows.length };
    });
    const issueIds = new Set(issues.map((issue) => String(issue.displayId || issue.id)));
    const predecessorStats = baseRows.reduce((acc, row) => {
      if (!row.predecessorFieldValue) acc.semPredecessor += 1;
      else if (row.predecessorIds.every((id) => issueIds.has(String(id)))) acc.predecessorValido += 1;
      else acc.predecessorNaoEncontrado += 1;
      return acc;
    }, { semPredecessor: 0, predecessorValido: 0, predecessorNaoEncontrado: 0 });
    console.info('[Modulo 8][Cronograma][Dashboard] Total de issues analisados:', baseRows.length);
    console.info('[Modulo 8][Cronograma][Dashboard] distribuição por status:', statusData);
    console.info('[Modulo 8][Cronograma][Dashboard] distribuição por tipo:', tipoData);
    console.info('[Modulo 8][Cronograma][Dashboard] distribuição por marco:', marcoData);
    console.info('[Modulo 8][Cronograma][Dashboard] distribuição por prazo:', prazoData);
    console.info('[Modulo 8][Cronograma][Dashboard] predecessores:', predecessorStats);
    return { statusData, marcoData, tipoData, prazoData, execucaoPorMarco, predecessorStats, total: baseRows.length };
  }, [cronogramaRows, issues]);

  const cronogramaProjectExecution = useMemo(() => {
    const baseRows = cronogramaRows.filter((row) => !row.isMarcoContratual);
    if (!baseRows.length) return 0;
    return Math.round(baseRows.reduce((sum, row) => sum + Number(row.executado || 0), 0) / baseRows.length);
  }, [cronogramaRows]);

  const eapUpcomingRows = useMemo(() => {
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const limit = new Date(today);
    limit.setDate(limit.getDate() + 7);
    return cronogramaRows
      .filter((row) => !row.concluded)
      .map((row) => ({ row, target: parseIssueDate(row.terminoPrevisto) || parseIssueDate(row.dataContratual) }))
      .filter((item) => item.target && item.target >= today && item.target <= limit)
      .sort((a, b) => a.target - b.target);
  }, [cronogramaRows]);

  const eapAttentionSummary = useMemo(() => ({
    atrasados: cronogramaRows.filter((row) => row.overdue && !row.concluded).length,
    bloqueados: cronogramaRows.filter((row) => row.blocked).length,
    semResponsavel: cronogramaRows.filter((row) => !row.atribuidoA || normalizeText(row.atribuidoA).includes('sem atribu')).length
  }), [cronogramaRows]);

  const eapDeliveryData = useMemo(() => {
    const deliveries = cronogramaRowsSorted.filter((row) => normalizeText(row.titulo).startsWith('entrega -'));
    const groups = deliveries.map((delivery) => {
      const prefix = String(delivery.eapVinculada || '').trim();
      const items = cronogramaRowsSorted.filter((row) => row.issue.id !== delivery.issue.id && prefix && String(row.eapVinculada || '').startsWith(`${prefix}.`));
      const activityItems = items.filter((row) => !row.isMarcoContratual);
      const progressBase = activityItems.length ? activityItems : [delivery];
      const progress = Math.round(progressBase.reduce((sum, row) => sum + Number(row.executado || 0), 0) / Math.max(1, progressBase.length));
      const overdue = items.filter((row) => row.overdue && !row.concluded).length;
      const blocked = items.filter((row) => row.blocked).length;
      return {
        id: delivery.issue.id,
        title: delivery.titulo,
        eap: delivery.eapVinculada,
        marco: delivery.marcoContratual,
        delivery,
        items,
        progress,
        total: items.length,
        overdue,
        blocked
      };
    });
    const linked = new Set(groups.flatMap((group) => [group.delivery.issue.id, ...group.items.map((row) => row.issue.id)]));
    const orphans = cronogramaRowsSorted.filter((row) => !linked.has(row.issue.id));
    return { groups, orphans };
  }, [cronogramaRowsSorted]);

  useEffect(() => {
    if (!eapDeliveryData.groups.length) {
      if (eapSelectedDeliveryId !== 'all') setEapSelectedDeliveryId('all');
      return;
    }
    if (eapSelectedDeliveryId !== 'all' && !eapDeliveryData.groups.some((group) => group.id === eapSelectedDeliveryId)) {
      setEapSelectedDeliveryId('all');
    }
  }, [eapDeliveryData, eapSelectedDeliveryId]);

  const eapSelectedDelivery = useMemo(() => {
    if (eapSelectedDeliveryId === 'all') return null;
    return eapDeliveryData.groups.find((group) => group.id === eapSelectedDeliveryId) || null;
  }, [eapDeliveryData, eapSelectedDeliveryId]);

  const eapQuickFilterCounts = useMemo(() => {
    const rows = eapSelectedDelivery ? [eapSelectedDelivery.delivery, ...eapSelectedDelivery.items] : cronogramaRowsSorted;
    return {
      all: rows.length,
      atrasadas: rows.filter((row) => row.overdue && !row.concluded).length,
      bloqueadas: rows.filter((row) => row.blocked).length,
      semResponsavel: rows.filter((row) => !row.atribuidoA || normalizeText(row.atribuidoA).includes('sem atribu')).length,
      vencem7: rows.filter((row) => {
        const target = parseIssueDate(row.terminoPrevisto) || parseIssueDate(row.dataContratual);
        if (!target || row.concluded) return false;
        const today = new Date();
        today.setHours(0,0,0,0);
        const limit = new Date(today);
        limit.setDate(limit.getDate() + 7);
        return target >= today && target <= limit;
      }).length,
      semPredecessor: rows.filter((row) => !String(row.predecessorFieldValue || '').trim()).length
    };
  }, [eapSelectedDelivery, cronogramaRowsSorted]);

  const applyEapQuickFilter = useCallback((rows) => rows.filter((row) => {
    if (eapQuickFilter === 'atrasadas') return row.overdue && !row.concluded;
    if (eapQuickFilter === 'bloqueadas') return row.blocked;
    if (eapQuickFilter === 'semResponsavel') return !row.atribuidoA || normalizeText(row.atribuidoA).includes('sem atribu');
    if (eapQuickFilter === 'vencem7') {
      const target = parseIssueDate(row.terminoPrevisto) || parseIssueDate(row.dataContratual);
      if (!target || row.concluded) return false;
      const today = new Date();
      today.setHours(0,0,0,0);
      const limit = new Date(today);
      limit.setDate(limit.getDate() + 7);
      return target >= today && target <= limit;
    }
    if (eapQuickFilter === 'semPredecessor') return !String(row.predecessorFieldValue || '').trim();
    return true;
  }), [eapQuickFilter]);

  const eapViewRows = useMemo(() => {
    let rows = [];
    if (eapStructureView === 'entregas') {
      rows = eapSelectedDelivery ? [eapSelectedDelivery.delivery, ...eapSelectedDelivery.items] : cronogramaRowsSorted;
    } else if (eapStructureView === 'cronograma') {
      rows = eapSelectedDelivery ? [eapSelectedDelivery.delivery, ...eapSelectedDelivery.items] : cronogramaRowsSorted;
      rows = [...rows].sort((a, b) => {
        const eapCompare = compareEapCodes(a.eapVinculada, b.eapVinculada);
        if (eapCompare) return eapCompare;
        const aDate = parseIssueDate(a.inicioPrevisto) || parseIssueDate(a.terminoPrevisto) || parseIssueDate(a.dataContratual) || new Date('2100-01-01');
        const bDate = parseIssueDate(b.inicioPrevisto) || parseIssueDate(b.terminoPrevisto) || parseIssueDate(b.dataContratual) || new Date('2100-01-01');
        return aDate - bDate;
      });
    } else if (eapStructureView === 'dependencias') {
      rows = (eapSelectedDelivery ? [eapSelectedDelivery.delivery, ...eapSelectedDelivery.items] : cronogramaRowsSorted)
        .filter((row) => row.blocked || String(row.predecessorFieldValue || '').trim() || row.predecessoresTexto.includes('ID não encontrado'));
    } else {
      rows = cronogramaRowsSorted;
    }
    return applyEapQuickFilter(rows);
  }, [eapStructureView, eapSelectedDelivery, eapDeliveryData, cronogramaRowsSorted, applyEapQuickFilter]);

  const eapMarcoSummary = useMemo(() => cronogramaDashboard.execucaoPorMarco.map((item) => ({
    ...item,
    progress: item.total ? Math.round((item.concluidos / item.total) * 100) : 0
  })), [cronogramaDashboard]);

  useEffect(() => {
    const candidateRows = eapStructureView === 'marcos' ? [] : eapViewRows;
    if (!candidateRows.length) {
      if (eapSelectedRowId) setEapSelectedRowId('');
      return;
    }
    if (!candidateRows.some((row) => row.issue.id === eapSelectedRowId)) {
      setEapSelectedRowId(candidateRows[0].issue.id);
    }
  }, [eapViewRows, eapSelectedRowId, eapStructureView]);

  const eapSelectedRow = useMemo(() => {
    const selected = eapViewRows.find((row) => row.issue.id === eapSelectedRowId);
    return selected || eapSelectedDelivery?.delivery || eapViewRows[0] || null;
  }, [eapViewRows, eapSelectedRowId, eapSelectedDelivery]);

  const eapTimelineData = useMemo(() => {
    const datedRows = eapViewRows.map((row) => {
      const plannedStart = parseIssueDate(row.inicioPrevisto) || parseIssueDate(row.startDate) || parseIssueDate(row.dataContratual) || parseIssueDate(row.dataPrevistaG5);
      const plannedEnd = parseIssueDate(row.terminoPrevisto) || parseIssueDate(row.dataPrevistaG5) || parseIssueDate(row.dataContratual) || plannedStart;
      const realStart = parseIssueDate(row.inicioReal);
      const realEnd = parseIssueDate(row.terminoReal) || (realStart ? new Date() : null);
      return { ...row, ganttStart: plannedStart, ganttEnd: plannedEnd || plannedStart, realStart, realEnd };
    });
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const validDates = datedRows.flatMap((row) => [row.ganttStart, row.ganttEnd, row.realStart, row.realEnd]).filter(Boolean);
    const dateTimes = [...validDates.map((date) => date.getTime()), today.getTime()];
    const minTime = dateTimes.length ? Math.min(...dateTimes) : today.getTime();
    const maxTime = dateTimes.length ? Math.max(...dateTimes) : today.getTime() + 86400000;
    const totalMs = Math.max(86400000, maxTime - minTime);
    const todayLeft = Math.max(0, Math.min(100, ((today.getTime() - minTime) / totalMs) * 100));
    const ticks = Array.from({ length: 6 }, (_, index) => new Date(minTime + (totalMs * index) / 5));
    const items = datedRows.map((row) => {
      const startMs = row.ganttStart?.getTime() || minTime;
      const endMs = row.ganttEnd?.getTime() || startMs + 86400000;
      const left = Math.max(0, Math.min(96, ((startMs - minTime) / totalMs) * 100));
      const width = Math.max(2.5, Math.min(100 - left, ((Math.max(endMs, startMs + 86400000) - startMs) / totalMs) * 100));
      const realStartMs = row.realStart?.getTime();
      const realEndMs = row.realEnd?.getTime();
      const realLeft = realStartMs ? Math.max(0, Math.min(96, ((realStartMs - minTime) / totalMs) * 100)) : 0;
      const realWidth = realStartMs
        ? Math.max(2.5, Math.min(100 - realLeft, ((Math.max(realEndMs || realStartMs + 86400000, realStartMs + 86400000) - realStartMs) / totalMs) * 100))
        : 0;
      return { ...row, ganttLeft: left, ganttWidth: width, realLeft, realWidth };
    });
    return { ticks, items, todayLeft };
  }, [eapViewRows]);

  const cronogramaGantt = useMemo(() => {
    const datedRows = cronogramaRowsFiltered.map((row) => {
      const editedRow = cronogramaEdits[row.issue.id] || {};
      const inicioPrevisto = editedRow.inicioPrevisto ?? row.inicioPrevisto;
      const tempoAtividade = editedRow.tempoAtividade ?? row.tempoAtividade;
      const terminoPrevisto = computeEndDateFromStartAndDays(inicioPrevisto, tempoAtividade) || editedRow.terminoPrevisto || row.terminoPrevisto;
      const inicioReal = editedRow.inicioReal ?? row.inicioReal;
      const terminoReal = editedRow.terminoReal ?? row.terminoReal;
      const plannedStart = parseIssueDate(inicioPrevisto) || parseIssueDate(row.startDate) || parseIssueDate(row.dataContratual) || parseIssueDate(row.dataPrevistaG5);
      const plannedEnd = parseIssueDate(terminoPrevisto) || parseIssueDate(row.dataPrevistaG5) || parseIssueDate(row.dataContratual) || plannedStart;
      const realStart = parseIssueDate(inicioReal);
      const realEnd = parseIssueDate(terminoReal) || (realStart ? new Date() : null);
      return {
        ...row,
        ganttStart: plannedStart,
        ganttEnd: plannedEnd || plannedStart,
        realStart,
        realEnd
      };
    });
    const validDates = datedRows.flatMap((row) => [row.ganttStart, row.ganttEnd, row.realStart, row.realEnd]).filter(Boolean);
    const minTime = validDates.length ? Math.min(...validDates.map((date) => date.getTime())) : Date.now();
    const maxTime = validDates.length ? Math.max(...validDates.map((date) => date.getTime())) : Date.now();
    const totalMs = Math.max(86400000, maxTime - minTime);
    const ticks = Array.from({ length: 6 }, (_, index) => new Date(minTime + (totalMs * index) / 5));
    const items = datedRows.map((row) => {
      const startMs = row.ganttStart?.getTime() || minTime;
      const endMs = row.ganttEnd?.getTime() || startMs + 86400000;
      const left = Math.max(0, Math.min(96, ((startMs - minTime) / totalMs) * 100));
      const width = Math.max(2.5, Math.min(100 - left, ((Math.max(endMs, startMs + 86400000) - startMs) / totalMs) * 100));
      const realStartMs = row.realStart?.getTime();
      const realEndMs = row.realEnd?.getTime();
      const realLeft = realStartMs ? Math.max(0, Math.min(96, ((realStartMs - minTime) / totalMs) * 100)) : 0;
      const realWidth = realStartMs
        ? Math.max(2.5, Math.min(100 - realLeft, ((Math.max(realEndMs || realStartMs + 86400000, realStartMs + 86400000) - realStartMs) / totalMs) * 100))
        : 0;
      return { ...row, ganttLeft: left, ganttWidth: width, realLeft, realWidth };
    });
    return { ticks, items };
  }, [cronogramaRowsFiltered, cronogramaEdits]);

  const scheduleFieldDefinitionsByKey = useMemo(() => {
    const entries = Object.entries(scheduleFieldAliases).map(([key, aliases]) => [key, findCustomFieldDefinition(customFieldDefinitions, aliases)]);
    return Object.fromEntries(entries);
  }, [customFieldDefinitions]);

  const currentUserProjectAccess = useMemo(() => {
    const currentUserKeys = new Set(
      [
        user?.id,
        user?.userId,
        user?.autodeskId,
        user?.uid,
        user?.email,
        user?.name,
        ...(Array.isArray(user?.ids) ? user.ids : [])
      ]
        .filter(Boolean)
        .map(normalizeText)
    );

    const matchedMember = projectUsers.find((projectUser) => {
      const memberKeys = [
        projectUser.id,
        projectUser.userId,
        projectUser.autodeskId,
        projectUser.uid,
        projectUser.email,
        projectUser.name
      ]
        .filter(Boolean)
        .map(normalizeText);

      return memberKeys.some((memberKey) => currentUserKeys.has(memberKey));
    });

    const accessText = normalizeText([
      matchedMember?.role,
      matchedMember?.projectRole,
      matchedMember?.accessLevel,
      matchedMember?.permissionLevel
    ].filter(Boolean).join(' '));

    return {
      member: matchedMember || null,
      isProjectAdmin: Boolean(
        matchedMember?.isProjectAdmin ||
        matchedMember?.isAdmin ||
        accessText.includes('admin') ||
        accessText.includes('administrador') ||
        accessText.includes('project admin') ||
        accessText.includes('account admin')
      )
    };
  }, [projectUsers, user]);

  const scheduleRows = useMemo(() => {
    const sourceIssueList = issues;
    const sourceIssuesById = new Map();
    sourceIssueList.forEach((issue) => {
      const uniqueKey = String(issue?.id || issue?.displayId || issue?.autodeskId || issue?.raw?.id || issue?.title || '').trim();
      if (!uniqueKey || sourceIssuesById.has(uniqueKey)) return;
      sourceIssuesById.set(uniqueKey, issue);
    });
    const sourceIssues = [...sourceIssuesById.values()];
    const lookup = buildIssueLookupIndexes(sourceIssues);
    const rawRows = sourceIssues.map((issue) => {
      const fields = Object.keys(scheduleFieldAliases).reduce((accumulator, key) => {
        accumulator[key] = getScheduleField(issue, customFieldDefinitions, key);
        return accumulator;
      }, {});
      const typeLabel = getIssueTypeLabel(issue);
      const normalizedType = normalizeText(typeLabel);
      const normalizedItemType = normalizeText(fields.tipoItemCronograma);
      const isMarco =
        normalizedType.includes('marco contratual') ||
        /\bmc\b/.test(normalizedType) ||
        normalizedItemType.includes('marco contratual') ||
        (normalizedItemType.includes('entrega tecnica') && (fields.codigoMarco || fields.dataContratual || fields.marcoContratual));
      const references = resolveIssueReferences(issue, lookup);
      const completed = isScheduleCompleted(issue, customFieldDefinitions, fields);
      const delay = calculateScheduleDelay(fields, completed);
      const predecessorIds = parsePredecessorIds(fields.predecessor);
      const predecessorIssues = predecessorIds
        .map((id) => lookup.byDisplayId.get(normalizeText(id)) || lookup.byId.get(String(id)) || lookup.byId.get(normalizeText(id)))
        .filter(Boolean);
      const predecessorMissing = predecessorIds.length > predecessorIssues.length;
      const predecessorPending = predecessorIssues.some((predecessor) => !isScheduleCompleted(predecessor, customFieldDefinitions));
      const issueProgress = getScheduleProgressFromDeliveryStatus(fields.statusEntrega);
      const impactText = fields.nivelImpacto || fields.impactoCronograma || fields.impactoMarco || fields.riscoPrevisto;
      return {
        id: issue.id,
        issue,
        fields,
        typeLabel,
        isMarco,
        completed,
        delay,
        references,
        referenceIds: new Set(references.flatMap((ref) => [ref.id, ref.displayId, ref.autodeskId].filter(Boolean).map((value) => normalizeText(value)))),
        predecessorIds,
        predecessorIssues,
        predecessorMissing,
        predecessorPending,
        issueProgress,
        impactText,
        critical: ['alto', 'critico', 'crítico'].some((item) => normalizeText(impactText).includes(item))
      };
    });

    const milestones = rawRows.filter((row) => row.isMarco);
    const childRows = rawRows.filter((row) => !row.isMarco);
    const membership = new Map(milestones.map((milestone) => [milestone.id, []]));
    const issueMatches = new Map();

    childRows.forEach((child) => {
      const matchGroups = [
        milestones.filter((milestone) => child.referenceIds.has(normalizeText(milestone.issue.id)) || child.referenceIds.has(normalizeText(milestone.issue.displayId))),
        milestones.filter((milestone) => child.fields.codigoMarco && normalizeText(child.fields.codigoMarco) === normalizeText(milestone.fields.codigoMarco)),
        milestones.filter((milestone) => child.fields.marcoContratual && normalizeText(child.fields.marcoContratual) === normalizeText(milestone.fields.marcoContratual)),
        milestones.filter((milestone) => child.fields.pacoteMarcoRelacionado && normalizeText(child.fields.pacoteMarcoRelacionado) === normalizeText(milestone.fields.pacoteMarcoRelacionado)),
        milestones.filter((milestone) => child.fields.eapVinculada && milestone.fields.eapVinculada && normalizeText(child.fields.eapVinculada).startsWith(normalizeText(milestone.fields.eapVinculada)))
      ].map((items) => items.filter((item, index, array) => array.findIndex((candidate) => candidate.id === item.id) === index));
      const bestMatch = matchGroups.find((items) => items.length) || [];
      issueMatches.set(child.id, bestMatch);
      const primaryMatch = bestMatch[0];
      if (primaryMatch) membership.get(primaryMatch.id)?.push(child);
    });

    const outputRows = [];
    milestones
      .sort((first, second) => String(first.fields.codigoMarco || first.fields.eapVinculada || first.issue.title).localeCompare(String(second.fields.codigoMarco || second.fields.eapVinculada || second.issue.title), 'pt-BR', { numeric: true, sensitivity: 'base' }))
      .forEach((milestone) => {
        const children = membership.get(milestone.id) || [];
        const completedChildren = children.filter((child) => child.completed).length;
        const delayedChildren = children.filter((child) => child.delay.days > 0).length;
        const criticalChildren = children.filter((child) => child.critical).length;
        const milestoneProgress = children.length
          ? Math.round(children.reduce((sum, child) => sum + (child.issueProgress || 0), 0) / children.length)
          : getScheduleProgressFromDeliveryStatus(milestone.fields.statusEntrega);
        const maxDelay = children.length ? Math.max(...children.map((child) => child.delay.days || 0)) : milestone.delay.days;
        const blocked = children.some((child) => child.predecessorPending || child.predecessorMissing || child.critical);
        const calculated = {
          avançoCalculado: milestoneProgress,
          diasAtrasoCalculado: maxDelay,
          situacaoPrazo: maxDelay > 0 ? 'Atrasado' : 'No prazo',
          filhosTotal: children.length,
          filhosConcluidos: completedChildren,
          filhosAtrasados: delayedChildren,
          filhosCriticos: criticalChildren,
          maiorAtraso: maxDelay,
          marcoBloqueado: blocked ? 'Sim' : 'Não',
          marcoAptoMedicao: milestoneProgress >= 100 ? 'Sim' : 'Não',
          statusExecutivo: blocked ? 'Atenção' : milestoneProgress >= 100 ? 'Concluído' : maxDelay > 0 ? 'Atrasado' : 'Em andamento'
        };
        outputRows.push({ ...milestone, level: 0, parentId: null, children, calculated, progress: milestoneProgress, maxDelay, linkWarning: children.length ? '' : 'Marco sem Issues vinculados' });
        children
          .sort((first, second) =>
            String(first.fields.predecessor || '').localeCompare(String(second.fields.predecessor || ''), 'pt-BR', { numeric: true, sensitivity: 'base' })
            || String(first.fields.codigoDocumentoCliente || first.fields.eapVinculada || first.issue.title).localeCompare(String(second.fields.codigoDocumentoCliente || second.fields.eapVinculada || second.issue.title), 'pt-BR', { numeric: true, sensitivity: 'base' })
          )
          .forEach((child) => {
            const matches = issueMatches.get(child.id) || [];
            outputRows.push({
              ...child,
              level: 1,
              parentId: milestone.id,
              children: [],
              calculated: {
                avançoCalculado: child.issueProgress,
                diasAtrasoCalculado: child.delay.days,
                situacaoPrazo: child.delay.days > 0 ? 'Atrasado' : 'No prazo',
                alertaPredecessor: child.predecessorPending ? 'Predecessor pendente' : child.predecessorMissing ? 'Predecessor não encontrado' : '',
                alertaVinculoMarco: matches.length > 1 ? 'Issue vinculado a mais de um marco' : ''
              },
              progress: child.issueProgress,
              maxDelay: child.delay.days,
              linkWarning: matches.length > 1 ? 'Issue vinculado a mais de um marco' : ''
            });
          });
      });

    const orphanRows = childRows.filter((child) => !(issueMatches.get(child.id) || []).length);
    orphanRows.forEach((orphan) => {
      outputRows.push({
        ...orphan,
        level: 1,
        parentId: null,
        children: [],
        calculated: {
          avançoCalculado: orphan.issueProgress,
          diasAtrasoCalculado: orphan.delay.days,
          situacaoPrazo: orphan.delay.days > 0 ? 'Atrasado' : 'No prazo',
          alertaVinculoMarco: 'Sem marco vinculado'
        },
        progress: orphan.issueProgress,
        maxDelay: orphan.delay.days,
        linkWarning: 'Sem marco vinculado'
      });
    });

    return outputRows;
  }, [issues, customFieldDefinitions]);

  const schedulePendingEditCount = useMemo(() => Object.keys(scheduleEdits).length, [scheduleEdits]);

  const scheduleFormulaSuggestions = useMemo(() => {
    const suggestions = [];
    const addSuggestion = (row, key, value, reason, safeToOverwrite = true) => {
      const definition = scheduleFieldDefinitionsByKey[key];
      if (!definition?.id || value === undefined || value === null || value === '') return;
      const current = row.fields[key] || '';
      if (normalizeText(current) === normalizeText(value)) return;
      suggestions.push({
        id: `${row.id}:${key}`,
        issueId: row.id,
        issueTitle: row.issue.title || row.issue.displayId || row.id,
        key,
        label: scheduleFieldLabels[key] || (definition.name || key),
        definition,
        current,
        value,
        reason,
        safeToOverwrite
      });
    };

    scheduleRows.forEach((row) => {
      const computedPercent = `${Math.round(row.progress || 0)}%`;
      if (row.isMarco) addSuggestion(row, 'percentualTecnico', computedPercent, 'Avanço calculado pelos Issues vinculados ao marco.');
      const deliveryStatus = inferScheduleDeliveryStatus(row.fields, row.completed, row.delay.days);
      addSuggestion(row, 'statusEntrega', deliveryStatus, 'Status consolidado por emissão, retorno e aprovação.');
      addSuggestion(row, 'statusCliente', inferScheduleClientStatus(row.fields, row.completed, row.delay.days), 'Status do cliente derivado da tramitação.');
      addSuggestion(row, 'impactoCronograma', scheduleImpactLevelFromDelay(row.delay.days), 'Impacto calculado pelos dias de atraso.');
      addSuggestion(row, 'riscoPrevisto', row.delay.days > 0 || row.predecessorPending || row.predecessorMissing ? scheduleImpactLevelFromDelay(row.delay.days || 4) : 'N.A.', 'Risco derivado de atraso, predecessor ou bloqueio.');
      addSuggestion(row, 'prioridadeGestao', schedulePriorityFromSignals({ delayDays: row.delay.days, impact: row.impactText, blocked: row.predecessorPending || row.predecessorMissing, critical: row.critical }), 'Prioridade sugerida por atraso, impacto e bloqueio.');
      if (!row.fields.acaoNecessaria) {
        const action = row.predecessorPending
          ? 'Validar predecessor pendente.'
          : row.delay.reason === 'Retorno do cliente'
            ? 'Cobrar retorno do cliente.'
            : row.delay.reason === 'Revisão interna'
              ? 'Concluir revisão interna.'
              : row.linkWarning
                ? 'Regularizar vínculo com marco.'
                : row.delay.days > 0
                  ? 'Verificar pendência de emissão.'
                  : '';
        addSuggestion(row, 'acaoNecessaria', action, 'Ação sugerida pelo motor de cronograma.', false);
      }
      if (!row.fields.dataLimiteRetornoCliente && row.fields.dataRealEmissao) {
        addSuggestion(row, 'dataLimiteRetornoCliente', formatBrazilianDate(addBusinessDays(parseScheduleDate(row.fields.dataRealEmissao), getBusinessDaysOrDefault(row.fields.prazoAnaliseCliente, 10))), 'Data Real de Emissão + prazo de análise do cliente.');
      }
      if (!row.fields.dataLimiteRevisaoInterna && row.fields.dataRealRetornoCliente) {
        addSuggestion(row, 'dataLimiteRevisaoInterna', formatBrazilianDate(addBusinessDays(parseScheduleDate(row.fields.dataRealRetornoCliente), getBusinessDaysOrDefault(row.fields.prazoRevisaoInterna, 5))), 'Data Real Retorno Cliente + prazo de revisão interna.');
      }
      const computedPlannedEnd = computeScheduleEndDateFromBusinessDays(
        scheduleEdits[row.id]?.inicioPlanejado ?? row.fields.inicioPlanejado,
        scheduleEdits[row.id]?.diasPrevistosAtividade ?? row.fields.diasPrevistosAtividade
      );
      if (computedPlannedEnd) {
        addSuggestion(row, 'terminoPlanejado', computedPlannedEnd, 'Inicio Planejado + Dias previstos para atividade em dias uteis.');
      }
    });
    return suggestions;
  }, [scheduleRows, scheduleFieldDefinitionsByKey, scheduleEdits]);

  const scheduleDashboard = useMemo(() => {
    const milestones = scheduleRows.filter((row) => row.isMarco);
    const linked = scheduleRows.filter((row) => !row.isMarco && row.parentId);
    const progressValues = scheduleRows.map((row) => Number(String(getScheduleCellValue(row, 'progress')).replace(/[^\d.-]/g, '')) || 0);
    const averageProgress = progressValues.length ? Math.round(progressValues.reduce((total, value) => total + value, 0) / progressValues.length) : 0;
    return {
      totalItems: scheduleRows.length,
      totalMarcos: milestones.length,
      totalVinculados: linked.length,
      marcosNoPrazo: milestones.filter((row) => !row.maxDelay).length,
      marcosAtrasados: milestones.filter((row) => row.maxDelay > 0).length,
      desenvolvimento: scheduleRows.filter((row) => normalizeText(row.fields.statusEntrega).includes('desenvolvimento') || normalizeText(row.fields.statusEntrega).includes('aguardando')).length,
      emitidas: scheduleRows.filter((row) => row.fields.dataRealEmissao || normalizeText(row.fields.statusEntrega).includes('emitid')).length,
      analiseCliente: scheduleRows.filter((row) => normalizeText(row.fields.statusCliente).includes('analise') || normalizeText(row.fields.statusEntrega).includes('analise pelo cliente')).length,
      comentarios: scheduleRows.filter((row) => normalizeText(row.fields.statusCliente).includes('coment')).length,
      revisaoInterna: scheduleRows.filter((row) => normalizeText(row.fields.statusEntrega).includes('revisao interna')).length,
      aprovadas: scheduleRows.filter((row) => row.completed).length,
      pendenciasAbertas: scheduleRows.filter((row) => normalizeText(row.typeLabel).includes('pendencia') && !row.completed).length,
      riscosAbertos: scheduleRows.filter((row) => (normalizeText(row.typeLabel).includes('risco') || normalizeText(row.typeLabel).includes('restricao')) && !row.completed).length,
      semMarco: scheduleRows.filter((row) => row.linkWarning === 'Sem marco vinculado').length,
      predecessorPendente: scheduleRows.filter((row) => row.predecessorPending || row.predecessorMissing).length,
      atualizacoesPendentes: scheduleFormulaSuggestions.length,
      alteracoesPendentes: schedulePendingEditCount,
      aguardandoCliente: scheduleRows.filter((row) => normalizeText(row.fields.statusEntrega).includes('cliente') || normalizeText(row.fields.statusCliente).includes('cliente')).length,
      avancoGeral: averageProgress
    };
  }, [scheduleRows, scheduleFormulaSuggestions.length, schedulePendingEditCount]);

  const scheduleStatusOptions = useMemo(() => {
    const values = scheduleRows
      .map((row) => String(scheduleEdits[row.id]?.statusEntrega ?? row.fields.statusEntrega ?? '').trim())
      .filter(Boolean);
    return Array.from(new Set(values)).sort((first, second) => first.localeCompare(second, 'pt-BR', { sensitivity: 'base' }));
  }, [scheduleRows, scheduleEdits]);

  const scheduleOwnerOptions = useMemo(() => {
    const values = scheduleRows
      .map((row) => getScheduleRowOwner(row))
      .filter(Boolean);
    return Array.from(new Set(values)).sort((first, second) => first.localeCompare(second, 'pt-BR', { sensitivity: 'base' }));
  }, [scheduleRows]);

  const scheduleMarcoOptions = useMemo(() => {
    const values = scheduleRows
      .map((row) => String(scheduleEdits[row.id]?.codigoMarco ?? row.fields.codigoMarco ?? '').trim())
      .filter(Boolean);
    return Array.from(new Set(values)).sort((first, second) => first.localeCompare(second, 'pt-BR', { numeric: true, sensitivity: 'base' }));
  }, [scheduleRows, scheduleEdits]);

  const scheduleKpiCards = useMemo(() => [
    { id: 'total', label: 'Itens no cronograma', value: scheduleDashboard.totalItems, tone: 'neutral' },
    { id: 'marcos', label: 'Marcos contratuais', value: scheduleDashboard.totalMarcos, tone: 'blue' },
    { id: 'tarefas', label: 'Tarefas vinculadas', value: scheduleDashboard.totalVinculados, tone: 'teal' },
    { id: 'atrasos', label: 'Em atraso', value: scheduleDashboard.marcosAtrasados, tone: scheduleDashboard.marcosAtrasados ? 'danger' : 'neutral' },
    { id: 'desenvolvimento', label: 'Em desenvolvimento', value: scheduleDashboard.desenvolvimento, tone: 'amber' },
    { id: 'cliente', label: 'Aguardando cliente', value: scheduleDashboard.aguardandoCliente, tone: 'purple' },
    { id: 'aprovadas', label: 'Aprovadas / concluidas', value: scheduleDashboard.aprovadas, tone: 'green' },
    { id: 'avanco', label: 'Avanco geral', value: `${scheduleDashboard.avancoGeral}%`, tone: 'progress' }
  ], [scheduleDashboard]);

  const filteredScheduleRows = useMemo(() => {
    const query = normalizeText(scheduleSearch);
    const valueForSort = (row, key) => String(scheduleEdits[row.id]?.[key] ?? row.fields[key] ?? '');
    return scheduleRows.filter((row) => {
      if (scheduleViewMode === 'critica') {
        const critical = row.maxDelay > 0 || row.critical || row.predecessorPending || row.predecessorMissing || row.linkWarning || !row.fields.dataContratual || !row.fields.dataMetaInterna || normalizeText(row.fields.statusCliente).includes('sem retorno');
        if (!critical) return false;
      }
      if (scheduleStatusFilter !== 'all' && normalizeText(valueForSort(row, 'statusEntrega')) !== normalizeText(scheduleStatusFilter)) return false;
      if (scheduleOwnerFilter !== 'all' && normalizeText(getScheduleRowOwner(row)) !== normalizeText(scheduleOwnerFilter)) return false;
      if (scheduleMarcoFilter !== 'all' && normalizeText(valueForSort(row, 'codigoMarco')) !== normalizeText(scheduleMarcoFilter)) return false;
      if (!query) return true;
      return [
        row.issue.title,
        row.issue.id,
        row.issue.displayId,
        row.fields.codigoMarco,
        row.fields.marcoContratual,
        row.fields.codigoDocumentoCliente,
        row.fields.codigoDocumentoInterno,
        row.fields.eapVinculada,
        row.fields.predecessor,
        row.fields.areaResponsavel,
        row.fields.disciplinaEnvolvida
      ].some((value) => normalizeText(value).includes(query));
    }).sort((first, second) =>
      valueForSort(first, 'codigoMarco').localeCompare(valueForSort(second, 'codigoMarco'), 'pt-BR', { numeric: true, sensitivity: 'base' })
      || (first.isMarco === second.isMarco ? 0 : first.isMarco ? -1 : 1)
      || valueForSort(first, 'predecessor').localeCompare(valueForSort(second, 'predecessor'), 'pt-BR', { numeric: true, sensitivity: 'base' })
      || String(first.issue.title || '').localeCompare(String(second.issue.title || ''), 'pt-BR', { numeric: true, sensitivity: 'base' })
    );
  }, [scheduleRows, scheduleSearch, scheduleViewMode, scheduleEdits, scheduleStatusFilter, scheduleOwnerFilter, scheduleMarcoFilter]);

  const schedulePlannerRows = useMemo(
    () => buildSchedulePlannerRows(filteredScheduleRows, schedulePlannerLayout),
    [filteredScheduleRows, schedulePlannerLayout]
  );

  const selectedScheduleRow = useMemo(
    () => schedulePlannerRows.find((row) => row.id === scheduleSelectedRowId) || null,
    [schedulePlannerRows, scheduleSelectedRowId]
  );

  const scheduleGantt = useMemo(() => {
    const rows = schedulePlannerRows.slice(0, 80).map((row) => {
      const plannedStartValue = getScheduleCellValue(row, 'inicioPlanejado');
      const computedPlannedEnd = computeScheduleEndDateFromBusinessDays(
        plannedStartValue,
        getScheduleCellValue(row, 'diasPrevistosAtividade')
      );
      const plannedStart = parseScheduleDate(plannedStartValue) || parseScheduleDate(row.fields.dataMetaInterna) || parseScheduleDate(row.fields.dataContratual);
      const plannedEndValue = computedPlannedEnd || getScheduleCellValue(row, 'terminoPlanejado');
      const plannedEnd = parseScheduleDate(plannedEndValue) || plannedStart;
      const realStart = parseScheduleDate(getScheduleCellValue(row, 'inicioReal'));
      const realEnd = parseScheduleDate(getScheduleCellValue(row, 'terminoReal')) || (realStart ? new Date() : null);
      return { ...row, plannedStart, plannedEnd, realStart, realEnd };
    });
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const dates = rows.flatMap((row) => [row.plannedStart, row.plannedEnd, row.realStart, row.realEnd, parseScheduleDate(row.fields.dataContratual), parseScheduleDate(row.fields.dataMetaInterna), today]).filter(Boolean);
    const minTime = dates.length ? Math.min(...dates.map((date) => date.getTime())) : today.getTime();
    const maxTime = dates.length ? Math.max(...dates.map((date) => date.getTime())) : addDays(today, 30).getTime();
    const totalMs = Math.max(86400000, maxTime - minTime);
    const pct = (date) => {
      const parsedDate = parseScheduleDate(date);
      if (!parsedDate) return null;
      return Math.max(0, Math.min(100, ((parsedDate.getTime() - minTime) / totalMs) * 100));
    };
    const ticks = Array.from({ length: 6 }, (_, index) => new Date(minTime + (totalMs * index) / 5));
    const items = rows.map((row) => {
      const start = row.plannedStart || row.plannedEnd || today;
      const end = row.plannedEnd || row.plannedStart || addDays(start, 1);
      const left = pct(start) ?? 0;
      const width = Math.max(2, (pct(end) ?? left + 2) - left);
      const realLeft = row.realStart ? pct(row.realStart) ?? 0 : 0;
      const realWidth = row.realStart ? Math.max(2, (pct(row.realEnd || today) ?? realLeft + 2) - realLeft) : 0;
      return {
        ...row,
        progress: Number(String(getScheduleCellValue(row, 'progress')).replace(/[^\d.-]/g, '')) || 0,
        ganttLeft: left,
        ganttWidth: Math.min(100 - left, width),
        realLeft,
        realWidth: Math.min(100 - realLeft, realWidth),
        markers: [
          { key: 'contratual', label: 'Data Contratual', color: '#1d74d8', left: pct(row.fields.dataContratual) },
          { key: 'meta', label: 'Data Meta Interna', color: '#2f9e44', left: pct(row.fields.dataMetaInterna) },
          { key: 'retorno', label: 'Limite Retorno Cliente', color: '#f08c00', left: pct(row.fields.dataLimiteRetornoCliente) },
          { key: 'retorno-real', label: 'Retorno Cliente', color: '#7048e8', left: pct(row.fields.dataRealRetornoCliente) },
          { key: 'aprovacao', label: 'Aprovação Final', color: '#111827', left: pct(row.fields.dataAprovacaoFinal) }
        ].filter((marker) => marker.left !== null)
      };
    });
    return { ticks, items, todayLeft: pct(today) };
  }, [schedulePlannerRows, scheduleEdits]);

  const scheduleColumns = useMemo(() => {
    const columns = {
      executiva: ['eapAuto', 'codigoMarco', 'title', 'inicioPlanejado', 'diasPrevistosAtividade', 'terminoPlanejado', 'inicioReal', 'terminoReal', 'dataContratual', 'dataLimiteInterna', 'predecessor', 'dependencia', 'statusEntrega', 'delay', 'acaoNecessaria', 'prioridadeGestao', 'dataPublicacao', 'progress'],
      coordenacao: ['eapAuto', 'codigoMarco', 'title', 'assignedTo', 'inicioPlanejado', 'diasPrevistosAtividade', 'terminoPlanejado', 'inicioReal', 'terminoReal', 'dataLimiteInterna', 'predecessor', 'dependencia', 'vinculoDependencia', 'statusEntrega', 'delay', 'acaoNecessaria', 'prioridadeGestao', 'dataPublicacao', 'progress'],
      tramitacao: ['eapAuto', 'codigoMarco', 'title', 'inicioPlanejado', 'terminoPlanejado', 'predecessor', 'dependencia', 'dataPublicacao', 'numeroTramitacao', 'statusCliente', 'dataRealEmissao', 'dataLimiteRetornoCliente', 'dataRealRetornoCliente', 'dataLimiteRevisaoInterna', 'dataRealReemissao', 'dataAprovacaoFinal'],
      critica: ['eapAuto', 'codigoMarco', 'title', 'inicioPlanejado', 'diasPrevistosAtividade', 'terminoPlanejado', 'inicioReal', 'terminoReal', 'dataContratual', 'dataLimiteInterna', 'predecessor', 'dependencia', 'impactoCronograma', 'prioridadeGestao', 'delay', 'acaoNecessaria'],
      completa: ['eapAuto', 'typeLabel', 'eapVinculada', 'codigoMarco', 'marcoContratual', 'codigoDocumentoCliente', 'codigoDocumentoInterno', 'title', 'disciplinaEnvolvida', 'areaResponsavel', 'fase', 'tipoItemCronograma', 'faseFluxo', 'statusEntrega', 'statusCliente', 'inicioPlanejado', 'diasPrevistosAtividade', 'terminoPlanejado', 'inicioReal', 'terminoReal', 'dataContratual', 'dataLimiteInterna', 'predecessor', 'dependencia', 'dataMetaInterna', 'dataPublicacao', 'dataRealEmissao', 'dataLimiteRetornoCliente', 'dataRealRetornoCliente', 'dataLimiteRevisaoInterna', 'dataAprovacaoFinal', 'impactoMarco', 'impactoCronograma', 'prioridadeGestao', 'percentualTecnico', 'progress', 'delay', 'acaoNecessaria']
    };
    return columns[scheduleViewMode] || columns.executiva;
  }, [scheduleViewMode]);

  const scheduleMissingCalculatedTargets = useMemo(() => {
    const wanted = ['percentualTecnico', 'statusEntrega', 'statusCliente', 'impactoCronograma', 'impactoMarco', 'riscoPrevisto', 'prioridadeGestao', 'acaoNecessaria', 'terminoPlanejado', 'dataLimiteRetornoCliente', 'dataLimiteRevisaoInterna', 'dataLimiteInterna'];
    return wanted.filter((key) => !scheduleFieldDefinitionsByKey[key]?.id).map((key) => scheduleFieldLabels[key] || scheduleFieldAliases[key]?.[0] || key);
  }, [scheduleFieldDefinitionsByKey]);

  const scheduleMissingStructuralTargets = useMemo(() => {
    const wanted = ['eapVinculada', 'predecessor'];
    return wanted.filter((key) => !scheduleFieldDefinitionsByKey[key]?.id).map((key) => scheduleFieldLabels[key] || scheduleFieldAliases[key]?.[0] || key);
  }, [scheduleFieldDefinitionsByKey]);

  const getEapLevelColor = useCallback((level) => {
    const palette = ['#f59e0b', '#22c55e', '#ef4444', '#a855f7', '#14b8a6', '#f97316', '#84cc16', '#e11d48'];
    const normalized = String(level || '').trim();
    if (!normalized) return '#9ca3af';
    const seed = normalized.split('').reduce((acc, ch) => acc + ch.charCodeAt(0), 0);
    return palette[seed % palette.length];
  }, []);

  function getScheduleStoredEap(row) {
    return String(row?.eapVinculada || row?.fields?.eapVinculada || '').trim();
  }

  function getScheduleEapParts(value) {
    const parts = String(value || '').trim().split('.').map((part) => part.trim()).filter(Boolean);
    if (!parts.length || parts.some((part) => !/^\d+$/.test(part))) return [];
    return parts;
  }

  function compareScheduleEapParts(firstParts, secondParts) {
    if (!firstParts.length && !secondParts.length) return 0;
    if (!firstParts.length) return 1;
    if (!secondParts.length) return -1;
    const length = Math.max(firstParts.length, secondParts.length);
    for (let index = 0; index < length; index += 1) {
      const firstValue = Number(firstParts[index] || 0);
      const secondValue = Number(secondParts[index] || 0);
      if (firstValue !== secondValue) return firstValue - secondValue;
    }
    return firstParts.length - secondParts.length;
  }

  function isScheduleCompletionValue(value) {
    const normalizedValue = normalizeText(value);
    return scheduleCompletionStatuses.some((status) => normalizedValue.includes(normalizeText(status)));
  }

  function buildSchedulePredecessorMaps(rows, edits = {}) {
    const rowById = new Map();
    const rowByLookupKey = new Map();
    rows.forEach((row) => {
      rowById.set(String(row.id), row);
      [
        row.id,
        row.issue?.id,
        row.issue?.displayId,
        row.issue?.autodeskId,
        getIssueHumanId(row.issue)
      ]
        .filter(Boolean)
        .forEach((value) => rowByLookupKey.set(normalizeText(String(value)), row));
    });

    const predecessorsByRowId = new Map();
    const missingByRowId = new Map();
    rows.forEach((row) => {
      const rawValue = normalizePredecessorInput(edits[row.id]?.predecessor ?? row.fields.predecessor);
      const predecessorIds = parsePredecessorIds(rawValue);
      const resolvedRows = [];
      const missingValues = [];
      predecessorIds.forEach((predecessorId) => {
        const matchedRow = rowByLookupKey.get(normalizeText(predecessorId));
        if (matchedRow) {
          resolvedRows.push(matchedRow);
        } else {
          missingValues.push(predecessorId);
        }
      });
      predecessorsByRowId.set(row.id, resolvedRows);
      missingByRowId.set(row.id, missingValues);
    });

    return { rowById, rowByLookupKey, predecessorsByRowId, missingByRowId };
  }

  function detectSchedulePredecessorCycle(rows, edits = {}) {
    const { predecessorsByRowId } = buildSchedulePredecessorMaps(rows, edits);
    const visiting = new Set();
    const visited = new Set();

    const visit = (rowId, trail = []) => {
      if (visiting.has(rowId)) return [...trail, rowId];
      if (visited.has(rowId)) return null;
      visiting.add(rowId);
      const predecessors = predecessorsByRowId.get(rowId) || [];
      for (const predecessorRow of predecessors) {
        if (!predecessorRow?.id) continue;
        const cycle = visit(predecessorRow.id, [...trail, rowId]);
        if (cycle) return cycle;
      }
      visiting.delete(rowId);
      visited.add(rowId);
      return null;
    };

    for (const row of rows) {
      const cycle = visit(row.id, []);
      if (cycle) return cycle;
    }
    return null;
  }

  function orderScheduleRowsByPredecessor(rows, edits = {}) {
    const baseRows = [...rows];
    const baseIndexById = new Map(baseRows.map((row, index) => [row.id, index]));
    const { predecessorsByRowId } = buildSchedulePredecessorMaps(baseRows, edits);
    const primaryPredecessorByRowId = new Map();
    const childrenByPredecessorId = new Map();

    baseRows.forEach((row) => {
      const primaryPredecessor = (predecessorsByRowId.get(row.id) || [])[0] || null;
      primaryPredecessorByRowId.set(row.id, primaryPredecessor?.id || null);
      if (!primaryPredecessor?.id) return;
      if (!childrenByPredecessorId.has(primaryPredecessor.id)) childrenByPredecessorId.set(primaryPredecessor.id, []);
      childrenByPredecessorId.get(primaryPredecessor.id).push(row);
    });

    childrenByPredecessorId.forEach((children) => {
      children.sort((first, second) => (baseIndexById.get(first.id) || 0) - (baseIndexById.get(second.id) || 0));
    });

    const ordered = [];
    const visited = new Set();
    const emit = (row) => {
      if (!row || visited.has(row.id)) return;
      visited.add(row.id);
      ordered.push(row);
      (childrenByPredecessorId.get(row.id) || []).forEach((child) => emit(child));
    };

    baseRows.filter((row) => !primaryPredecessorByRowId.get(row.id)).forEach((row) => emit(row));
    baseRows.forEach((row) => emit(row));
    return ordered;
  }

  function formatScheduleCycleMessage(cycleIds, rows) {
    const rowById = new Map(rows.map((row) => [row.id, row]));
    const labels = cycleIds.map((rowId) => {
      const row = rowById.get(rowId);
      return row ? (getIssueHumanId(row.issue) || row.issue?.title || rowId) : rowId;
    });
    return labels.join(' → ');
  }

  function applyScheduleOrderFromDependencies(nextEdits) {
    const orderedRows = orderScheduleRowsByPredecessor(scheduleRows, nextEdits);
    setSchedulePlannerLayout((current) => Object.fromEntries(
      orderedRows.map((row, index) => {
        const currentLayout = current[row.id] || {};
        const fallbackLevel = Number(row.plannerLevel ?? row.level ?? 0);
        return [
          row.id,
          {
            order: index,
            level: Number.isFinite(Number(currentLayout.level)) ? Number(currentLayout.level) : fallbackLevel,
            parentId: currentLayout.parentId ?? row.plannerParentId ?? row.parentId ?? null
          }
        ];
      })
    ));
  }

  function buildSchedulePredecessorValidation(issueId, value, currentEdits) {
    const currentIssueEdits = currentEdits[issueId] || {};
    const nextEdits = {
      ...currentEdits,
      [issueId]: {
        ...currentIssueEdits,
        predecessor: value
      }
    };
    const currentRow = scheduleRows.find((row) => row.id === issueId);
    if (!currentRow) return { valid: true, nextEdits };

    const { rowByLookupKey, missingByRowId } = buildSchedulePredecessorMaps(scheduleRows, nextEdits);
    const missing = missingByRowId.get(issueId) || [];
    if (missing.length) {
      return {
        valid: false,
        message: `Predecessor não encontrado: ${missing.join(', ')}. Verifique o ID do issue antes de salvar.`,
        nextEdits
      };
    }

    const directPredecessors = parsePredecessorIds(value);
    const selfReference = directPredecessors.some((predecessorId) => {
      const matchedRow = rowByLookupKey.get(normalizeText(predecessorId));
      return matchedRow?.id === issueId;
    });
    if (selfReference) {
      return {
        valid: false,
        message: 'Um item não pode apontar para si mesmo como predecessor.',
        nextEdits
      };
    }

    const cycle = detectSchedulePredecessorCycle(scheduleRows, nextEdits);
    if (cycle) {
      return {
        valid: false,
        message: `Alteração bloqueada: o predecessor informado cria um ciclo de dependência (${formatScheduleCycleMessage(cycle, scheduleRows)}).`,
        nextEdits
      };
    }

    return { valid: true, nextEdits };
  }

  function resolveSchedulePredecessorRows(row, edits = scheduleEdits) {
    const rawValue = normalizePredecessorInput(edits[row.id]?.predecessor ?? row.fields.predecessor);
    const predecessorIds = parsePredecessorIds(rawValue);
    const lookup = new Map();
    scheduleRows.forEach((candidate) => {
      [
        candidate.id,
        candidate.issue?.id,
        candidate.issue?.displayId,
        candidate.issue?.autodeskId,
        getIssueHumanId(candidate.issue)
      ]
        .filter(Boolean)
        .forEach((value) => lookup.set(normalizeText(String(value)), candidate));
    });
    return predecessorIds.map((predecessorId) => ({
      id: predecessorId,
      row: lookup.get(normalizeText(predecessorId)) || null
    }));
  }

  function getScheduleDependencyState(row, edits = scheduleEdits) {
    const predecessorLinks = resolveSchedulePredecessorRows(row, edits);
    const dependencyValue = edits[row.id]?.dependencia ?? row.localDependency ?? row.fields.dependencia ?? 'Nao';
    const dependencyEnabled = normalizeText(dependencyValue).startsWith('s');

    if (!predecessorLinks.length) {
      return {
        status: 'none',
        label: 'Sem predecessor',
        detail: '',
        blocked: false,
        released: false,
        missing: false,
        enabled: dependencyEnabled,
        predecessorLinks
      };
    }

    const missing = predecessorLinks.filter((link) => !link.row);
    if (missing.length) {
      return {
        status: 'missing',
        label: '⚠️ Predecessor não encontrado',
        detail: missing.map((link) => link.id).join(', '),
        blocked: true,
        released: false,
        missing: true,
        enabled: dependencyEnabled,
        predecessorLinks
      };
    }

    if (!dependencyEnabled) {
      const firstLink = predecessorLinks[0];
      return {
        status: 'info',
        label: `🔗 Vínculo informativo ${getIssueHumanId(firstLink.row.issue) || firstLink.id}`,
        detail: firstLink.row.issue?.title || '',
        blocked: false,
        released: false,
        missing: false,
        enabled: dependencyEnabled,
        predecessorLinks
      };
    }

    const pending = predecessorLinks.filter((link) => !isScheduleCompleted(link.row.issue, customFieldDefinitions, {
      ...link.row.fields,
      ...edits[link.row.id]
    }));
    if (pending.length) {
      const firstPending = pending[0];
      return {
        status: 'blocked',
        label: `🔒 Bloqueada por ${getIssueHumanId(firstPending.row.issue) || firstPending.id}`,
        detail: firstPending.row.issue?.title || '',
        blocked: true,
        released: false,
        missing: false,
        enabled: dependencyEnabled,
        predecessorLinks
      };
    }

    const firstLink = predecessorLinks[0];
    return {
      status: 'released',
      label: `✅ Liberada por ${getIssueHumanId(firstLink.row.issue) || firstLink.id}`,
      detail: firstLink.row.issue?.title || '',
      blocked: false,
      released: true,
      missing: false,
      enabled: dependencyEnabled,
      predecessorLinks
    };
  }

  function validateScheduleEapChange(issueId, value, currentEdits) {
    const normalizedValue = String(value || '').trim();
    if (normalizedValue && !/^\d+(\.\d*)*$/.test(normalizedValue)) {
      return {
        valid: false,
        message: 'EAP inválida. Use somente números separados por ponto, por exemplo: 2.3.7.1.'
      };
    }
    if (!normalizedValue || normalizedValue.endsWith('.')) return { valid: true };

    const duplicatedRow = scheduleRows.find((row) => {
      if (row.id === issueId) return false;
      const candidateValue = String(currentEdits[row.id]?.eapVinculada ?? row.fields.eapVinculada ?? '').trim();
      return candidateValue && normalizeText(candidateValue) === normalizeText(normalizedValue);
    });
    if (duplicatedRow) {
      return {
        valid: false,
        message: `EAP duplicada. O código ${normalizedValue} já está vinculado ao item ${duplicatedRow.issue?.title || getIssueHumanId(duplicatedRow.issue)}.`
      };
    }

    return { valid: true };
  }

  function applyScheduleOrderFromEap(nextEdits) {
    const orderedRows = [...scheduleRows].sort((first, second) => {
      const firstValue = nextEdits[first.id]?.eapVinculada ?? first.fields.eapVinculada ?? '';
      const secondValue = nextEdits[second.id]?.eapVinculada ?? second.fields.eapVinculada ?? '';
      const eapCompare = compareScheduleEapParts(getScheduleEapParts(firstValue), getScheduleEapParts(secondValue));
      if (eapCompare) return eapCompare;
      return String(first.issue?.title || '').localeCompare(String(second.issue?.title || ''), 'pt-BR', { numeric: true, sensitivity: 'base' });
    });

    setSchedulePlannerLayout((current) => Object.fromEntries(
      orderedRows.map((row, index) => {
        const currentLayout = current[row.id] || {};
        const eapParts = getScheduleEapParts(nextEdits[row.id]?.eapVinculada ?? row.fields.eapVinculada ?? '');
        return [
          row.id,
          {
            order: index,
            level: eapParts.length ? Math.max(0, eapParts.length - 1) : (Number.isFinite(Number(currentLayout.level)) ? Number(currentLayout.level) : row.level || 0),
            parentId: currentLayout.parentId ?? row.plannerParentId ?? row.parentId ?? null
          }
        ];
      })
    ));
  }

  function buildSchedulePlannerRows(rows, layout) {
    const hasSavedLayout = Object.keys(layout || {}).length > 0;
    const baseIndexById = new Map(rows.map((row, index) => [row.id, index]));
    const idByStoredEap = new Map();
    rows.forEach((row) => {
      const storedParts = getScheduleEapParts(getScheduleStoredEap(row));
      if (storedParts.length) idByStoredEap.set(storedParts.join('.'), row.id);
    });
    const orderedRows = [...rows].sort((first, second) => {
      const firstOrder = layout[first.id]?.order;
      const secondOrder = layout[second.id]?.order;
      if (Number.isFinite(firstOrder) && Number.isFinite(secondOrder)) return firstOrder - secondOrder;
      if (Number.isFinite(firstOrder)) return firstOrder - baseIndexById.get(second.id);
      if (Number.isFinite(secondOrder)) return baseIndexById.get(first.id) - secondOrder;
      const storedCompare = compareScheduleEapParts(
        getScheduleEapParts(getScheduleStoredEap(first)),
        getScheduleEapParts(getScheduleStoredEap(second))
      );
      if (storedCompare) return storedCompare;
      return baseIndexById.get(first.id) - baseIndexById.get(second.id);
    });
    const counters = [];
    return orderedRows.map((row, index) => {
      const storedEapParts = getScheduleEapParts(getScheduleStoredEap(row));
      const useStoredEap = !hasSavedLayout && storedEapParts.length > 0;
      const savedLevel = Number(layout[row.id]?.level);
      const baseLevel = Number(row.plannerLevel ?? (useStoredEap ? storedEapParts.length - 1 : row.level ?? 0));
      const plannerLevel = Math.max(0, Math.min(6, Number.isFinite(savedLevel) ? savedLevel : baseLevel));
      counters[plannerLevel] = (counters[plannerLevel] || 0) + 1;
      counters.length = plannerLevel + 1;
      for (let levelIndex = 0; levelIndex < plannerLevel; levelIndex += 1) {
        if (!counters[levelIndex]) counters[levelIndex] = 1;
      }
      const generatedEap = counters.slice(0, plannerLevel + 1).join('.');
      const eapAuto = useStoredEap ? storedEapParts.join('.') : generatedEap;
      const storedParentEap = useStoredEap ? storedEapParts.slice(0, -1).join('.') : '';
      const storedParentId = storedParentEap ? idByStoredEap.get(storedParentEap) || null : null;
      return {
        ...row,
        plannerOrder: index,
        plannerLevel,
        plannerParentId: layout[row.id]?.parentId ?? row.plannerParentId ?? storedParentId ?? row.parentId ?? null,
        eapAuto
      };
    });
  }

  function commitSchedulePlannerRows(nextRows, extraChangesById = {}) {
    const transientLayout = Object.fromEntries(
      nextRows.map((row, index) => [
        row.id,
        {
          order: index,
          level: row.plannerLevel || 0,
          parentId: row.plannerParentId || null
        }
      ])
    );
    const normalizedRows = buildSchedulePlannerRows(nextRows, transientLayout);
    const rowById = new Map(normalizedRows.map((row) => [row.id, row]));
    setSchedulePlannerLayout(Object.fromEntries(
      normalizedRows.map((row, index) => [
        row.id,
        {
          order: index,
          level: row.plannerLevel,
          parentId: row.plannerParentId || null
        }
      ])
    ));
    setScheduleEdits((current) => {
      const next = { ...current };
      normalizedRows.forEach((row) => {
        const parentRow = row.plannerParentId ? rowById.get(row.plannerParentId) : null;
        const parentPredecessor = parentRow ? getIssueHumanId(parentRow.issue) : '';
        next[row.id] = {
          ...(next[row.id] || {}),
          eapVinculada: row.eapAuto,
          predecessor: parentPredecessor,
          ...(extraChangesById[row.id] || {})
        };
      });
      return next;
    });
  }

  function getSelectedScheduleIndex() {
    return schedulePlannerRows.findIndex((row) => row.id === scheduleSelectedRowId);
  }

  function selectScheduleRow(rowId) {
    setScheduleSelectedRowId(rowId);
  }

  function getScheduleHighlightStorageKey(projectId = selectedProjectId) {
    return `central-g5-schedule-row-highlights:${projectId || 'sem-projeto'}`;
  }

  function toggleSelectedScheduleDeliverableTitle() {
    if (!scheduleSelectedRowId || !selectedProjectId) return;
    setScheduleRowHighlights((current) => {
      const next = { ...current };
      if (next[scheduleSelectedRowId] === 'deliverableTitle') {
        delete next[scheduleSelectedRowId];
      } else {
        next[scheduleSelectedRowId] = 'deliverableTitle';
      }
      localStorage.setItem(getScheduleHighlightStorageKey(), JSON.stringify(next));
      return next;
    });
  }

  function moveSelectedScheduleRow(direction) {
    const selectedIndex = getSelectedScheduleIndex();
    const targetIndex = selectedIndex + direction;
    if (selectedIndex < 0 || targetIndex < 0 || targetIndex >= schedulePlannerRows.length) return;
    const nextRows = [...schedulePlannerRows];
    [nextRows[selectedIndex], nextRows[targetIndex]] = [nextRows[targetIndex], nextRows[selectedIndex]];
    commitSchedulePlannerRows(nextRows);
    setScheduleSelectedRowId(nextRows[targetIndex].id);
  }

  function indentSelectedScheduleRow() {
    const selectedIndex = getSelectedScheduleIndex();
    if (selectedIndex <= 0) return;
    const previousRow = schedulePlannerRows[selectedIndex - 1];
    const selectedLevel = schedulePlannerRows[selectedIndex].plannerLevel || 0;
    let blockEnd = selectedIndex + 1;
    while (blockEnd < schedulePlannerRows.length && (schedulePlannerRows[blockEnd].plannerLevel || 0) > selectedLevel) {
      blockEnd += 1;
    }
    const nextRows = schedulePlannerRows.map((row, index) => {
      if (index < selectedIndex || index >= blockEnd) return row;
      const isSelectedRow = index === selectedIndex;
      return {
        ...row,
        plannerLevel: Math.min(6, (row.plannerLevel || 0) + 1),
        plannerParentId: isSelectedRow ? previousRow.id : row.plannerParentId
      };
    });
    const predecessorValue = getIssueHumanId(previousRow.issue);
    commitSchedulePlannerRows(nextRows, {
      [schedulePlannerRows[selectedIndex].id]: {
        predecessor: predecessorValue,
        codigoMarco: previousRow.isMarco ? (previousRow.fields.codigoMarco || previousRow.eapAuto || '') : schedulePlannerRows[selectedIndex].fields.codigoMarco
      }
    });
  }

  function outdentSelectedScheduleRow() {
    const selectedIndex = getSelectedScheduleIndex();
    if (selectedIndex < 0) return;
    const selectedRow = schedulePlannerRows[selectedIndex];
    if ((selectedRow.plannerLevel || 0) <= 0) return;
    const nextLevel = Math.max(0, (selectedRow.plannerLevel || 0) - 1);
    const parentCandidate = [...schedulePlannerRows.slice(0, selectedIndex)].reverse().find((row) => (row.plannerLevel || 0) < nextLevel);
    let blockEnd = selectedIndex + 1;
    while (blockEnd < schedulePlannerRows.length && (schedulePlannerRows[blockEnd].plannerLevel || 0) > (selectedRow.plannerLevel || 0)) {
      blockEnd += 1;
    }
    const nextRows = schedulePlannerRows.map((row, index) => {
      if (index < selectedIndex || index >= blockEnd) return row;
      const isSelectedRow = index === selectedIndex;
      return {
        ...row,
        plannerLevel: Math.max(0, (row.plannerLevel || 0) - 1),
        plannerParentId: isSelectedRow ? parentCandidate?.id || null : row.plannerParentId
      };
    });
    commitSchedulePlannerRows(nextRows);
  }

  function normalizeIssueTypeMatch(value) {
    return normalizeText(value).replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function findScheduleIssueType(categoryName, typeName) {
    const wantedCategory = normalizeIssueTypeMatch(categoryName);
    const wantedType = normalizeIssueTypeMatch(typeName);
    const matchesWanted = (value, wanted) => {
      const normalizedValue = normalizeIssueTypeMatch(value);
      return normalizedValue === wanted || normalizedValue.includes(wanted) || wanted.includes(normalizedValue);
    };
    const candidates = issueTypeOptions.filter((item) => {
      const categoryText = item.category || item.typeTitle || item.title;
      return matchesWanted(categoryText, wantedCategory);
    });
    return (
      candidates.find((item) => item.kind === 'subtype' && matchesWanted(item.title, wantedType)) ||
      candidates.find((item) => matchesWanted(item.title, wantedType)) ||
      null
    );
  }

  async function createScheduleIssue(config) {
    if (!selectedProjectId) return;
    const selectedType = findScheduleIssueType(config.category, config.type);
    if (!selectedType?.id) {
      setError(`Nao encontrei no ACC o tipo "${config.type}" dentro da categoria "${config.category}". Confira se ele esta ativo nas configuracoes do projeto.`);
      return;
    }
    const title = window.prompt('Informe o titulo do novo item do cronograma:', config.defaultTitle);
    if (!title) return;
    const selectedRow = schedulePlannerRows.find((row) => row.id === scheduleSelectedRowId);
    const customAttributes = [];
    const pushCustom = (key, value) => {
      const definition = scheduleFieldDefinitionsByKey[key];
      if (definition?.id && value !== undefined && value !== null && value !== '') {
        customAttributes.push({ attributeDefinitionId: definition.id, value });
      }
    };
    if (config.usePlanningFields) {
      pushCustom('eapVinculada', String(schedulePlannerRows.length + 1));
    }
    if (selectedRow && config.usePlanningFields && !config.isMarco) {
      pushCustom('codigoMarco', selectedRow.fields.codigoMarco || selectedRow.eapAuto || '');
      pushCustom('predecessor', getIssueHumanId(selectedRow.issue));
    }
    setError('');
    try {
      const createdIssue = await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues`, {
        method: 'POST',
        body: JSON.stringify({
          title,
          description: config.description,
          issueTypeId: selectedType.kind === 'type' ? selectedType.id : undefined,
          issueSubtypeId: selectedType.kind === 'subtype' ? selectedType.id : undefined,
          customAttributes
        })
      });
      const createdIssueId = createdIssue?.id || createdIssue?.issue?.id || createdIssue?.data?.id;
      await refreshCurrentModule();
      if (createdIssueId) {
        setScheduleSelectedRowId(createdIssueId);
      }
    } catch (requestError) {
      setError(requestError.message);
    }
  }

  useEffect(() => {
    if (!selectedProjectId) {
      setKanbanManualColumns([]);
      setKanbanColumnMessage('');
      return;
    }

    const storageKey = `g5-kanban-manual-marcos:${selectedProjectId}`;
    try {
      const storedColumns = JSON.parse(localStorage.getItem(storageKey) || '[]');
      setKanbanManualColumns(Array.isArray(storedColumns) ? storedColumns : []);
    } catch {
      setKanbanManualColumns([]);
    }
    setKanbanColumnMessage('');
  }, [selectedProjectId]);

  function updateKanbanDraft(field, value) {
    setKanbanDraft((current) => ({ ...current, [field]: value }));
  }

  function updateKanbanEditDraft(field, value) {
    setKanbanEditDraft((current) => ({ ...current, [field]: value }));
  }

  function normalizeKanbanColumnKey(value) {
    return normalizeText(value).replace(/[^a-z0-9]+/g, ' ').trim();
  }

  function getKanbanManualColumnsStorageKey(projectId = selectedProjectId) {
    return projectId ? `g5-kanban-manual-marcos:${projectId}` : '';
  }

  function persistKanbanManualColumns(projectId, columns) {
    const storageKey = getKanbanManualColumnsStorageKey(projectId);
    if (!storageKey) return;
    try {
      localStorage.setItem(storageKey, JSON.stringify(columns));
    } catch {
      // Mantem a coluna em memoria mesmo se o navegador bloquear o localStorage.
    }
  }

  function updateKanbanColumnDraft(field, value) {
    setKanbanColumnDraft((current) => ({ ...current, [field]: value }));
  }

  function createKanbanMarcoColumn(event) {
    event.preventDefault();
    const title = String(kanbanColumnDraft.title || '').trim();
    const code = String(kanbanColumnDraft.code || '').trim();

    if (!title) {
      setError('Informe o nome do Marco Contratual para criar a coluna.');
      return;
    }

    const titleKey = normalizeKanbanColumnKey(title);
    const codeKey = normalizeKanbanColumnKey(code);
    const duplicated = linkKanbanColumns.some((column) => {
      const columnTitleKey = normalizeKanbanColumnKey(getKanbanColumnLinkLabel(column));
      const columnCodeKey = normalizeKanbanColumnKey(column.code);
      return columnTitleKey === titleKey || (codeKey && columnCodeKey === codeKey);
    });

    if (duplicated) {
      setError('Ja existe uma coluna com este Marco Contratual ou Código do Marco.');
      return;
    }

    const manualColumn = {
      id: `manual-marco-${Date.now()}-${titleKey || 'marco'}`,
      code,
      title,
      deliveryTitle: title,
      marcoContratual: title,
      marcoFieldValue: title,
      canReceiveCards: true,
      isManual: true
    };

    setKanbanManualColumns((current) => {
      const nextColumns = [...current, manualColumn];
      persistKanbanManualColumns(selectedProjectId, nextColumns);
      return nextColumns;
    });
    setKanbanColumnDraft({ title: '', code: '' });
    setKanbanColumnMessage('Coluna de Marco Contratual criada apenas no quadro. Ao mover cards para ela, o campo Marco Contratual será atualizado no ACC.');
    setError('');
  }

  function removeKanbanManualColumn(columnId) {
    const column = linkKanbanColumns.find((item) => item.id === columnId);
    if (!column?.isManual) return;
    if (column.cards?.length) {
      setError('Esta coluna possui cards. Mova os cards antes de remover a coluna.');
      return;
    }
    setKanbanManualColumns((current) => {
      const nextColumns = current.filter((item) => item.id !== columnId);
      persistKanbanManualColumns(selectedProjectId, nextColumns);
      return nextColumns;
    });
    setKanbanColumnMessage('Coluna manual removida do quadro.');
  }

  function getKanbanAssignedUserId(issue) {
    const rawAssigned = issue?.assignedToId || issue?.raw?.assignedTo || issue?.raw?.assignedToId || issue?.raw?.assignedToUserId || '';
    if (rawAssigned && projectUsers.some((projectUser) => String(projectUser.id) === String(rawAssigned))) {
      return String(rawAssigned);
    }

    const assignedLabel = normalizeText(issue?.assignedTo || rawAssigned || '');
    if (!assignedLabel) return '';

    const matchedUser = projectUsers.find((projectUser) => {
      const possibleValues = [
        projectUser.id,
        projectUser.userId,
        projectUser.autodeskId,
        projectUser.name,
        projectUser.email
      ];
      return possibleValues.some((value) => {
        if (!value) return false;
        return String(value) === String(rawAssigned) || normalizeText(value) === assignedLabel;
      });
    });

    return matchedUser?.id || '';
  }

  function getKanbanProjectUserId(userRef) {
    if (!userRef) return '';
    const rawValue = typeof userRef === 'object'
      ? userRef.id || userRef.userId || userRef.autodeskId || userRef.autodesk_id || userRef.uid || userRef.email || userRef.name
      : userRef;
    if (!rawValue) return '';

    const normalizedRaw = normalizeText(rawValue);
    const matchedUser = projectUsers.find((projectUser) => {
      const possibleValues = [
        projectUser.id,
        projectUser.userId,
        projectUser.autodeskId,
        projectUser.autodesk_id,
        projectUser.uid,
        projectUser.name,
        projectUser.email
      ];
      return possibleValues.some((value) => value && (String(value) === String(rawValue) || normalizeText(value) === normalizedRaw));
    });

    return matchedUser?.id || '';
  }

  function getKanbanFollowerIds(issue) {
    const possibleLists = [
      issue?.followers,
      issue?.watchers,
      issue?.raw?.followers,
      issue?.raw?.watchers,
      issue?.raw?.followerIds,
      issue?.raw?.watcherIds
    ];
    const followerIds = new Set();

    possibleLists.forEach((list) => {
      const normalizedList = Array.isArray(list) ? list : list ? [list] : [];
      normalizedList.forEach((userRef) => {
        const userId = getKanbanProjectUserId(userRef);
        if (userId) followerIds.add(String(userId));
      });
    });

    return Array.from(followerIds);
  }

  function haveSameKanbanUserList(first = [], second = []) {
    const normalizeList = (list) => [...new Set((list || []).filter(Boolean).map(String))].sort();
    const normalizedFirst = normalizeList(first);
    const normalizedSecond = normalizeList(second);
    return normalizedFirst.length === normalizedSecond.length && normalizedFirst.every((value, index) => value === normalizedSecond[index]);
  }

  function getKanbanDisciplineOptionValue(issue) {
    const discipline = getIssueDiscipline(issue);
    if (!discipline) return '';
    const normalizedDiscipline = normalizeText(discipline);
    const matchedOption = kanbanDisciplineOptions.find((option) =>
      normalizeText(option.id) === normalizedDiscipline || normalizeText(option.label) === normalizedDiscipline
    );
    return matchedOption?.id || discipline;
  }

  function getKanbanMarcoTitleField() {
    return findCustomFieldDefinition(customFieldDefinitions, KANBAN_MARCO_TITLE_FIELD_ALIASES);
  }

  function getKanbanMarcoContratualField() {
    return findCustomFieldDefinition(customFieldDefinitions, MARCO_CONTRATUAL_FIELD_ALIASES);
  }

  function getKanbanColumnLinkLabel(column) {
    if (!column) return '';
    const deliveryTitle = String(column.deliveryTitle || '').trim();
    const title = String(column.title || '').trim();
    const marcoContratual = String(column.marcoContratual || '').trim();
    const code = String(column.code || '').trim();
    return deliveryTitle || title || marcoContratual || code;
  }

  function getKanbanColumnOptionLabel(column) {
    if (!column) return '';
    const code = String(column.code || '').trim();
    const mainLabel = getKanbanColumnLinkLabel(column) || 'Entrega sem nome';
    if (!code) return mainLabel;
    return normalizeText(mainLabel).includes(normalizeText(code)) ? mainLabel : `${mainLabel} | ${code}`;
  }

  function getKanbanColumnSubtitle(column) {
    if (!column) return '';
    const code = String(column.code || '').trim();
    if (!code || normalizeText(code) === 'sem marco') return '';
    return `Código do Marco: ${code}`;
  }

  function getKanbanWriteValue(field, preferredValues) {
    const candidates = preferredValues.map((value) => String(value || '').trim()).filter(Boolean);
    if (!candidates.length) return '';
    return resolveCustomFieldOptionValue(field, candidates[0]);
  }

  function mergeIssueCustomAttributeValues(issue, updates) {
    const validUpdates = updates.filter((update) => update?.field?.id && update.value !== undefined && update.value !== null);
    if (!validUpdates.length) return issue.customAttributes || [];

    const nextCustomAttributes = [...(issue.customAttributes || [])];
    validUpdates.forEach(({ field, value }) => {
      const fieldId = String(field.id || field.attributeDefinitionId || '');
      let foundAttribute = false;
      for (let index = 0; index < nextCustomAttributes.length; index += 1) {
        const attribute = nextCustomAttributes[index];
        const attributeId = String(attribute.id || attribute.attributeDefinitionId || '');
        const sameField =
          attributeId === fieldId ||
          normalizeFieldKey(attribute.name) === normalizeFieldKey(field.name || field.title || field.displayName);
        if (!sameField) continue;
        foundAttribute = true;
        nextCustomAttributes[index] = { ...attribute, value, rawValue: value, displayValue: value };
      }
      if (!foundAttribute) {
        nextCustomAttributes.push({
          id: field.id || field.attributeDefinitionId,
          attributeDefinitionId: field.id || field.attributeDefinitionId,
          name: field.name || field.title || field.displayName,
          value,
          rawValue: value,
          displayValue: value
        });
      }
    });
    return nextCustomAttributes;
  }

  function buildKanbanCustomAttributes(marcoCode, disciplinaValue, marcoTitle) {
    const customAttributes = [];
    const codigoMarcoField = findCustomFieldDefinition(customFieldDefinitions, CODIGO_MARCO_FIELD_ALIASES);
    const marcoContratualField = getKanbanMarcoContratualField();
    const marcoTitleField = getKanbanMarcoTitleField();
    const disciplinaField = findCustomFieldDefinition(customFieldDefinitions, ['Disciplina envolvida', 'Disciplina']);
    const marcoContratualValue = getKanbanWriteValue(marcoContratualField, [marcoTitle, marcoCode]);
    const marcoTitleValue = getKanbanWriteValue(marcoTitleField, [marcoTitle]);

    if (codigoMarcoField?.id && marcoCode) {
      customAttributes.push({ attributeDefinitionId: codigoMarcoField.id, value: marcoCode });
    }
    if (marcoContratualField?.id && marcoContratualValue && String(marcoContratualField.id) !== String(codigoMarcoField?.id || '')) {
      customAttributes.push({ attributeDefinitionId: marcoContratualField.id, value: marcoContratualValue });
    }
    if (
      marcoTitleField?.id &&
      marcoTitleValue &&
      String(marcoTitleField.id) !== String(codigoMarcoField?.id || '') &&
      String(marcoTitleField.id) !== String(marcoContratualField?.id || '')
    ) {
      customAttributes.push({ attributeDefinitionId: marcoTitleField.id, value: marcoTitleValue });
    }
    if (disciplinaField?.id && disciplinaValue) {
      customAttributes.push({ attributeDefinitionId: disciplinaField.id, value: disciplinaValue });
    }

    return customAttributes;
  }

  async function createKanbanCard(event) {
    event.preventDefault();
    if (!selectedProjectId) return;

    const selectedColumn = linkKanbanColumns.find((column) => column.id === kanbanDraft.marcoId);
    if (!selectedColumn?.canReceiveCards) {
      setError('Selecione um marco contratual para criar o card.');
      return;
    }
    if (!kanbanDraft.type || !kanbanDraft.title || !kanbanDraft.assignedTo || !kanbanDraft.dueDate || !kanbanDraft.disciplina) {
      setError('Preencha marco, tipo, título, responsável, data prevista e disciplina antes de criar o card.');
      return;
    }

    const selectedType = findScheduleIssueType('Interface e Coordenação Multidisciplinar', kanbanDraft.type);
    if (!selectedType?.id) {
      setError(`Nao encontrei no ACC o tipo "${kanbanDraft.type}" dentro da categoria "Interface e Coordenação Multidisciplinar".`);
      return;
    }

    setKanbanCreating(true);
    setError('');

    try {
      await requestJson('/api/auth/keep-alive', { method: 'POST' });
      await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues`, {
        method: 'POST',
        body: JSON.stringify({
          title: kanbanDraft.title,
          description: kanbanDraft.description,
          dueDate: kanbanDraft.dueDate,
          assignedTo: kanbanDraft.assignedTo,
          issueTypeId: selectedType.kind === 'type' ? selectedType.id : undefined,
          issueSubtypeId: selectedType.kind === 'subtype' ? selectedType.id : undefined,
          customAttributes: buildKanbanCustomAttributes(
            selectedColumn.code,
            kanbanDraft.disciplina,
            getKanbanColumnLinkLabel(selectedColumn)
          )
        })
      });

      setKanbanDraft((current) => ({
        ...current,
        title: '',
        dueDate: '',
        description: ''
      }));
      await refreshCurrentModule();
      setActiveModule('links');
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setKanbanCreating(false);
    }
  }

  function buildKanbanMoveUpdates(column) {
    const codigoMarcoField = findCustomFieldDefinition(customFieldDefinitions, CODIGO_MARCO_FIELD_ALIASES);
    const marcoContratualField = getKanbanMarcoContratualField();
    const marcoTitleField = getKanbanMarcoTitleField();

    if (!marcoContratualField?.id) {
      throw new Error('Nao encontrei o campo personalizado "Marco Contratual" para atualizar o card no ACC.');
    }

    const columnTitle = getKanbanColumnLinkLabel(column);
    const marcoContratualValue = getKanbanWriteValue(marcoContratualField, [
      columnTitle,
      column.title,
      column.marcoContratual,
      column.code
    ]);
    const marcoTitleValue = getKanbanWriteValue(marcoTitleField, [
      columnTitle,
      column.title,
      column.marcoContratual,
      column.code
    ]);

    const updates = [];
    if (marcoContratualField?.id && marcoContratualValue) {
      updates.push({ field: marcoContratualField, value: marcoContratualValue });
    }
    if (codigoMarcoField?.id && column.code && String(codigoMarcoField.id) !== String(marcoContratualField.id)) {
      updates.push({ field: codigoMarcoField, value: column.code });
    }
    if (
      marcoTitleField?.id &&
      marcoTitleValue &&
      String(marcoTitleField.id) !== String(marcoContratualField.id) &&
      String(marcoTitleField.id) !== String(codigoMarcoField?.id || '')
    ) {
      updates.push({ field: marcoTitleField, value: marcoTitleValue });
    }

    if (!updates.length) {
      throw new Error('Nao foi possivel preparar a atualizacao do Marco Contratual para este card.');
    }

    return updates;
  }

  async function moveKanbanCard(issue, nextMarcoId) {
    const nextColumn = linkKanbanColumns.find((column) => column.id === nextMarcoId);
    if (!issue?.id || !nextColumn?.canReceiveCards) return;

    const previousCustomAttributes = issue.customAttributes || [];
    let updates = [];

    try {
      updates = buildKanbanMoveUpdates(nextColumn);
    } catch (requestError) {
      setError(requestError.message);
      return;
    }

    setKanbanMovingIssueId(issue.id);
    setSelectedKanbanIssueId(issue.id);
    setError('');

    setIssues((currentIssues) =>
      currentIssues.map((currentIssue) =>
        currentIssue.id === issue.id
          ? {
              ...currentIssue,
              customAttributes: mergeIssueCustomAttributeValues(currentIssue, updates)
            }
          : currentIssue
      )
    );

    try {
      await requestJson('/api/auth/keep-alive', { method: 'POST' });
      const customAttributes = updates.map(({ field, value }) => {
        const prepared = prepareCustomFieldWriteValue(field, value);
        if (!prepared.valid) {
          throw new Error(`O valor "${value}" nao existe nas opcoes do campo "${field.name || field.title || field.displayName || 'personalizado'}" neste projeto do ACC.`);
        }
        return {
          attributeDefinitionId: field.id || field.attributeDefinitionId,
          value: prepared.value
        };
      });

      const updatedIssue = await requestJson(
        `/api/projects/${encodeURIComponent(selectedProjectId)}/issues/${encodeURIComponent(issue.id)}`,
        {
          method: 'PATCH',
          body: JSON.stringify({ customAttributes })
        }
      );

      setIssues((currentIssues) =>
        currentIssues.map((currentIssue) =>
          currentIssue.id === issue.id ? { ...currentIssue, ...updatedIssue } : currentIssue
        )
      );
      setKanbanColumnMessage(`Marco Contratual atualizado automaticamente no ACC para "${getKanbanColumnLinkLabel(nextColumn)}".`);
      setKanbanMoveDrafts((current) => {
        const nextDrafts = { ...current };
        delete nextDrafts[issue.id];
        return nextDrafts;
      });
      await refreshCurrentModule();
      setActiveModule('links');
    } catch (requestError) {
      setIssues((currentIssues) =>
        currentIssues.map((currentIssue) =>
          currentIssue.id === issue.id ? { ...currentIssue, customAttributes: previousCustomAttributes } : currentIssue
        )
      );
      setError(requestError.message);
    } finally {
      setKanbanMovingIssueId('');
    }
  }

  async function saveKanbanMoves() {
    const draftEntries = Object.entries(kanbanMoveDrafts);
    if (!draftEntries.length) {
      setKanbanColumnMessage('Os movimentos do Kanban agora são salvos automaticamente no ACC ao soltar o card na coluna.');
      return;
    }

    setKanbanSavingMoves(true);
    try {
      for (const [issueId, draft] of draftEntries) {
        const issue = issues.find((item) => item.id === issueId);
        if (issue) await moveKanbanCard(issue, draft.columnId);
      }
      setKanbanMoveDrafts({});
    } finally {
      setKanbanSavingMoves(false);
    }
  }

  function handleKanbanDragStart(event, issue) {
    if (!issue?.id) return;
    setKanbanDraggedIssueId(issue.id);
    setSelectedKanbanIssueId(issue.id);
    event.dataTransfer.effectAllowed = 'move';
    event.dataTransfer.setData('text/plain', issue.id);
  }

  function handleKanbanDragEnd() {
    setKanbanDraggedIssueId('');
    setKanbanDragOverColumnId('');
  }

  async function handleKanbanDrop(event, column) {
    event.preventDefault();
    const issueId = event.dataTransfer.getData('text/plain') || kanbanDraggedIssueId;
    setKanbanDragOverColumnId('');
    setKanbanDraggedIssueId('');

    if (!column?.canReceiveCards || !issueId) return;
    const issue = issues.find((item) => item.id === issueId);
    if (!issue) return;

    await moveKanbanCard(issue, column.id);
    setSelectedKanbanIssueId(issue.id);
  }

  async function saveKanbanCardEdit(event) {
    event.preventDefault();
    if (!selectedKanbanIssue?.id) return;

    const selectedColumn = linkKanbanColumns.find((column) => column.id === kanbanEditDraft.marcoId);
    const codigoMarcoField = findCustomFieldDefinition(customFieldDefinitions, CODIGO_MARCO_FIELD_ALIASES);
    const marcoContratualField = getKanbanMarcoContratualField();
    const marcoTitleField = getKanbanMarcoTitleField();
    const marcoContratualValue = getKanbanWriteValue(marcoContratualField, [
      getKanbanColumnLinkLabel(selectedColumn),
      selectedColumn?.title,
      selectedColumn?.code
    ]);
    const marcoTitleValue = getKanbanWriteValue(marcoTitleField, [
      getKanbanColumnLinkLabel(selectedColumn),
      selectedColumn?.title,
      selectedColumn?.code
    ]);
    const disciplinaField = findCustomFieldDefinition(customFieldDefinitions, ['Disciplina envolvida', 'Disciplina']);
    const nextDescription = kanbanEditDraft.comment
      ? `${kanbanEditDraft.description || ''}\n\nComentário registrado no Kanban em ${formatDateTime(new Date().toISOString())}:\n${kanbanEditDraft.comment}`.trim()
      : kanbanEditDraft.description;

    const patch = {};
    if (kanbanEditDraft.title !== (selectedKanbanIssue.title || '')) patch.title = kanbanEditDraft.title;
    if (nextDescription !== (selectedKanbanIssue.description || '')) patch.description = nextDescription;
    if (kanbanEditDraft.dueDate !== getDateInputValue(selectedKanbanIssue.dueDate)) patch.dueDate = kanbanEditDraft.dueDate || null;
    if (kanbanEditDraft.assignedTo !== getKanbanAssignedUserId(selectedKanbanIssue)) {
      patch.assignedTo = kanbanEditDraft.assignedTo || null;
      if (kanbanEditDraft.assignedTo) patch.assignedToType = 'user';
    }
    if (!haveSameKanbanUserList(kanbanEditDraft.followers, getKanbanFollowerIds(selectedKanbanIssue))) {
      patch.watchers = kanbanEditDraft.followers || [];
    }

    setKanbanSavingIssueId(selectedKanbanIssue.id);
    setError('');

    try {
      await requestJson('/api/auth/keep-alive', { method: 'POST' });
      if (Object.keys(patch).length) {
        await updateIssue(selectedKanbanIssue.id, patch);
      }
      if (selectedColumn?.canReceiveCards && codigoMarcoField?.id && selectedColumn.code !== getCustomFieldValue(selectedKanbanIssue, customFieldDefinitions, 'Código do Marco')) {
        await updateSingleCustomAttribute(selectedKanbanIssue, codigoMarcoField, selectedColumn.code);
      }
      if (selectedColumn?.canReceiveCards && marcoContratualField?.id && String(marcoContratualField.id) !== String(codigoMarcoField?.id || '') && marcoContratualValue) {
        const currentMarcoTitle = getCustomFieldValueByAliases(selectedKanbanIssue, customFieldDefinitions, MARCO_CONTRATUAL_FIELD_ALIASES);
        if (normalizeText(currentMarcoTitle) !== normalizeText(marcoContratualValue)) {
          await updateSingleCustomAttribute(selectedKanbanIssue, marcoContratualField, marcoContratualValue);
        }
      }
      if (
        selectedColumn?.canReceiveCards &&
        marcoTitleField?.id &&
        String(marcoTitleField.id) !== String(codigoMarcoField?.id || '') &&
        String(marcoTitleField.id) !== String(marcoContratualField?.id || '') &&
        marcoTitleValue
      ) {
        const currentMarcoTitle = getCustomFieldValueByAliases(selectedKanbanIssue, customFieldDefinitions, KANBAN_MARCO_TITLE_FIELD_ALIASES);
        if (normalizeText(currentMarcoTitle) !== normalizeText(marcoTitleValue)) {
          await updateSingleCustomAttribute(selectedKanbanIssue, marcoTitleField, marcoTitleValue);
        }
      }
      if (disciplinaField?.id && kanbanEditDraft.disciplina !== getKanbanDisciplineOptionValue(selectedKanbanIssue)) {
        await updateSingleCustomAttribute(selectedKanbanIssue, disciplinaField, kanbanEditDraft.disciplina || '');
      }

      setKanbanEditDraft((current) => ({ ...current, comment: '' }));
      await refreshCurrentModule();
      setSelectedKanbanIssueId(selectedKanbanIssue.id);
      setActiveModule('links');
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setKanbanSavingIssueId('');
    }
  }

  useEffect(() => {
    setCronogramaCurrentPage(1);
  }, [selectedProjectId, activeModule]);

  useEffect(() => {
    setSchedulePlannerLayout({});
    setScheduleSelectedRowId('');
    if (!selectedProjectId) {
      setScheduleRowHighlights({});
      return;
    }
    try {
      const savedHighlights = JSON.parse(localStorage.getItem(getScheduleHighlightStorageKey(selectedProjectId)) || '{}');
      setScheduleRowHighlights(savedHighlights && typeof savedHighlights === 'object' ? savedHighlights : {});
    } catch {
      setScheduleRowHighlights({});
    }
  }, [selectedProjectId]);

  useEffect(() => {
    if (scheduleSelectedRowId && !schedulePlannerRows.some((row) => row.id === scheduleSelectedRowId)) {
      setScheduleSelectedRowId('');
    }
  }, [schedulePlannerRows, scheduleSelectedRowId]);

  useEffect(() => {
    return () => {
      if (scheduleAutosaveTimerRef.current) window.clearTimeout(scheduleAutosaveTimerRef.current);
    };
  }, []);

  useEffect(() => {
    if (activeModule !== 'schedule' || !selectedProjectId) return undefined;
    if (!Object.keys(scheduleEdits).length) {
      if (!savingSchedule && scheduleAutosaveState !== 'error') {
        setScheduleAutosaveState('idle');
        setScheduleAutosaveMessage('');
      }
      return undefined;
    }

    setScheduleAutosaveState((current) => (current === 'saving' ? current : 'pending'));
    setScheduleAutosaveMessage((current) => current || 'Alterações pendentes. Salvamento automático em andamento.');
    if (scheduleAutosaveTimerRef.current) window.clearTimeout(scheduleAutosaveTimerRef.current);
    scheduleAutosaveTimerRef.current = window.setTimeout(() => {
      saveScheduleChanges({ silent: true, includeSuggestions: false });
    }, 1200);

    return () => {
      if (scheduleAutosaveTimerRef.current) window.clearTimeout(scheduleAutosaveTimerRef.current);
    };
  }, [activeModule, selectedProjectId, scheduleEdits]);

  useEffect(() => {
    if (cronogramaCurrentPage > cronogramaTotalPages) {
      setCronogramaCurrentPage(cronogramaTotalPages);
    }
  }, [cronogramaCurrentPage, cronogramaTotalPages]);

  function editCronogramaCell(issueId, field, value) {
    setCronogramaEdits((current) => ({
      ...current,
      [issueId]: {
        ...(current[issueId] || {}),
        [field]: value
      }
    }));
  }

  async function saveCronogramaChanges() {
    if (!selectedProjectId || !cronogramaPendingChangesCount) return;
    setSavingCronograma(true);
    setError('');
    try {
      const findCronogramaFieldDefinition = (aliases) =>
        (customFieldDefinitions || []).find((field) => {
          const fieldName = normalizeFieldKey(field.name || field.title || field.displayName || field.label);
          return aliases.some((alias) => fieldName === normalizeFieldKey(alias));
        });
      const dataContratualDef = findCronogramaFieldDefinition(['Data Contratual', 'Data contratual']);
      const dataPrevistaDef = findCronogramaFieldDefinition(['Data Prevista G5', 'Data prevista G5']);
      const inicioPrevistoDef = findCronogramaFieldDefinition(['Inicio Previsto', 'Início Previsto', 'Data Inicio Previsto', 'Data Início Previsto']);
      const terminoPrevistoDef = findCronogramaFieldDefinition(['Termino Previsto', 'Término Previsto', 'Data Termino Previsto', 'Data Término Previsto']);
      const inicioRealDef = findCronogramaFieldDefinition(['Inicio Real', 'Início Real', 'Data Inicio Real', 'Data Início Real']);
      const terminoRealDef = findCronogramaFieldDefinition(['Termino Real', 'Término Real', 'Data Termino Real', 'Data Término Real']);
      const predecessorDef = findCronogramaFieldDefinition(['Predecessor']);
      const tempoAtividadeDef = findCronogramaFieldDefinition(['Dias previstos para atividade', 'Dias Previsto para Atividade', 'Dias Previstos para Atividade', 'Dias previstos atividade', 'Dias previstos', 'Tempo para Atividade', 'Tempo para atividade']);
      for (const [issueId, changes] of Object.entries(cronogramaEdits)) {
        const currentRow = cronogramaRows.find((row) => row.issue.id === issueId);
        const computedTerminoPrevisto = computeEndDateFromStartAndDays(
          changes.inicioPrevisto ?? currentRow?.inicioPrevisto,
          changes.tempoAtividade ?? currentRow?.tempoAtividade
        );
        const customAttributes = [];
        if (changes.dataContratual !== undefined && dataContratualDef?.id) customAttributes.push({ attributeDefinitionId: dataContratualDef.id, value: changes.dataContratual || '' });
        if (changes.dataPrevistaG5 !== undefined && dataPrevistaDef?.id) customAttributes.push({ attributeDefinitionId: dataPrevistaDef.id, value: changes.dataPrevistaG5 || '' });
        if (changes.inicioPrevisto !== undefined && inicioPrevistoDef?.id) customAttributes.push({ attributeDefinitionId: inicioPrevistoDef.id, value: changes.inicioPrevisto || '' });
        if ((changes.terminoPrevisto !== undefined || changes.inicioPrevisto !== undefined || changes.tempoAtividade !== undefined) && terminoPrevistoDef?.id) customAttributes.push({ attributeDefinitionId: terminoPrevistoDef.id, value: computedTerminoPrevisto || changes.terminoPrevisto || '' });
        if (changes.inicioReal !== undefined && inicioRealDef?.id) customAttributes.push({ attributeDefinitionId: inicioRealDef.id, value: changes.inicioReal || '' });
        if (changes.terminoReal !== undefined && terminoRealDef?.id) customAttributes.push({ attributeDefinitionId: terminoRealDef.id, value: changes.terminoReal || '' });
        if (changes.predecessor !== undefined && predecessorDef?.id) customAttributes.push({ attributeDefinitionId: predecessorDef.id, value: normalizePredecessorInput(changes.predecessor) });
        if (changes.tempoAtividade !== undefined && tempoAtividadeDef?.id) {
          const tempoValue = String(changes.tempoAtividade || '').trim().replace(',', '.');
          const numericTempo = Number(tempoValue);
          customAttributes.push({
            attributeDefinitionId: tempoAtividadeDef.id,
            value: tempoValue === '' ? '' : Number.isFinite(numericTempo) ? numericTempo : tempoValue
          });
        }
        const payload = { customAttributes };
        if (changes.status !== undefined) payload.status = changes.status;
        if (!customAttributes.length && payload.status === undefined) continue;
        console.info('[Modulo 8][Cronograma] Payload enviado ao ACC', { issueId, payload });
        const updatedIssue = await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues/${encodeURIComponent(issueId)}`, { method: 'PATCH', body: JSON.stringify(payload) });
        console.info('[Modulo 8][Cronograma] Resposta ACC', { issueId, updatedIssue });
        setIssues((current) => current.map((issue) => (issue.id === issueId ? { ...issue, ...updatedIssue } : issue)));
      }
      setCronogramaEdits({});
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setSavingCronograma(false);
    }
  }

  function editScheduleCell(issueId, field, value) {
    const rowForEdit = scheduleRows.find((item) => item.id === issueId);

    if (field === 'predecessor') {
      const validation = buildSchedulePredecessorValidation(issueId, value, scheduleEdits);
      if (!validation.valid) {
        window.alert(validation.message);
        return;
      }
      setScheduleEdits(validation.nextEdits);
      applyScheduleOrderFromDependencies(validation.nextEdits);
      return;
    }

    if (field === 'eapVinculada') {
      const nextEdits = {
        ...scheduleEdits,
        [issueId]: {
          ...(scheduleEdits[issueId] || {}),
          eapVinculada: value
        }
      };
      const validation = validateScheduleEapChange(issueId, value, nextEdits);
      if (!validation.valid) {
        window.alert(validation.message);
        return;
      }
      setScheduleEdits(nextEdits);
      applyScheduleOrderFromEap(nextEdits);
      return;
    }

    if (field === 'statusEntrega' && rowForEdit) {
      const dependencyState = getScheduleDependencyState(rowForEdit, scheduleEdits);
      const nextProgress = getScheduleProgressFromDeliveryStatus(value);
      const isWaitingStatus = normalizeText(value).includes('aguard') || normalizeText(value).includes('pendente') || !value;
      if (dependencyState.blocked && nextProgress > 0 && !isWaitingStatus) {
        window.alert(`${dependencyState.label}. Conclua o predecessor antes de iniciar esta tarefa.`);
        return;
      }
    }

    setScheduleEdits((current) => {
      let nextIssueEdits = {
        ...(current[issueId] || {}),
        [field]: value
      };

      if (field === 'statusEntrega') {
        const row = rowForEdit || scheduleRows.find((item) => item.id === issueId);
        const previousStatus = current[issueId]?.statusEntrega ?? row?.fields?.statusEntrega ?? '';
        const statusChanged = normalizeText(previousStatus) !== normalizeText(value);
        const currentInicioReal = current[issueId]?.inicioReal ?? row?.fields?.inicioReal ?? '';
        const currentTerminoReal = current[issueId]?.terminoReal ?? row?.fields?.terminoReal ?? '';
        const todayValue = formatInputDate(new Date());

        if (statusChanged && !currentInicioReal) {
          nextIssueEdits.inicioReal = todayValue;
        }
        if (statusChanged && isScheduleCompletionValue(value) && !currentTerminoReal) {
          nextIssueEdits.terminoReal = todayValue;
        }
      }

      return {
        ...current,
        [issueId]: nextIssueEdits
      };
    });

    if (field === 'statusEntrega') {
      setScheduleAutosaveState('pending');
      setScheduleAutosaveMessage('Alteração de status pendente de salvamento automático.');
    }
  }

  function editScheduleDependency(row, value) {
    const predecessorIssue = row.predecessorIssues?.[0];
    const nextChanges = { dependencia: value };
    if (normalizeText(value).startsWith('s') && predecessorIssue) {
      const predecessorStart = getScheduleField(predecessorIssue, customFieldDefinitions, 'inicioPlanejado');
      const predecessorDays = getScheduleField(predecessorIssue, customFieldDefinitions, 'diasPrevistosAtividade');
      const predecessorPlannedEnd = getScheduleField(predecessorIssue, customFieldDefinitions, 'terminoPlanejado')
        || computeScheduleEndDateFromBusinessDays(predecessorStart, predecessorDays)
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataMetaInterna')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataContratual');
      const predecessorRealEnd = getScheduleField(predecessorIssue, customFieldDefinitions, 'terminoReal')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataAprovacaoFinal')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataRealReemissao')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataRealRetornoCliente');
      if (parseScheduleDate(predecessorPlannedEnd)) {
        nextChanges.inicioPlanejado = formatBrazilianDate(addBusinessDays(parseScheduleDate(predecessorPlannedEnd), 1));
      }
      if (parseScheduleDate(predecessorRealEnd)) {
        nextChanges.inicioReal = formatBrazilianDate(addBusinessDays(parseScheduleDate(predecessorRealEnd), 1));
      }
    }
    setScheduleEdits((current) => ({
      ...current,
      [row.id]: {
        ...(current[row.id] || {}),
        ...nextChanges
      }
    }));
  }

  function toggleScheduleFormulaSuggestion(suggestionId, checked) {
    setScheduleFormulaSelections((current) => ({ ...current, [suggestionId]: checked }));
  }

  function selectAllScheduleFormulaSuggestions(checked) {
    setScheduleFormulaSelections(
      Object.fromEntries(scheduleFormulaSuggestions.map((suggestion) => [suggestion.id, checked]))
    );
  }

  function getScheduleSaveValue(key, value) {
    if (['dataContratual', 'dataLimiteInterna', 'dataMetaInterna', 'dataPublicacao', 'inicioPlanejado', 'terminoPlanejado', 'dataRealEmissao', 'dataLimiteRetornoCliente', 'dataRealRetornoCliente', 'dataLimiteRevisaoInterna', 'dataRealReemissao', 'dataAprovacaoFinal', 'inicioReal', 'terminoReal'].includes(key)) {
      return formatScheduleDateForSave(value);
    }
    if (key === 'percentualTecnico') {
      const numericValue = Number(String(value || '').replace(',', '.').replace(/[^\d.-]/g, ''));
      return Number.isFinite(numericValue) ? `${Math.round(numericValue)}%` : value;
    }
    if (key === 'diasPrevistosAtividade') {
      const numericValue = Number(String(value || '').replace(',', '.').replace(/[^\d.-]/g, ''));
      return Number.isFinite(numericValue) ? numericValue : value;
    }
    return value ?? '';
  }

  async function saveScheduleChanges(options = {}) {
    if (!selectedProjectId) return;
    const { silent = false, includeSuggestions = !silent } = options;
    const selectedSuggestions = includeSuggestions
      ? scheduleFormulaSuggestions.filter((suggestion) => scheduleFormulaSelections[suggestion.id] !== false)
      : [];
    const accScheduleEdits = Object.fromEntries(
      Object.entries(scheduleEdits)
        .map(([issueId, changes]) => [
          issueId,
          Object.fromEntries(Object.entries(changes).filter(() => true))
        ])
        .filter(([, changes]) => Object.keys(changes).length)
    );
    const manualChangeCount = Object.keys(accScheduleEdits).length;
    if (!selectedSuggestions.length && !manualChangeCount) return;

    const structuralChanges = ['eapVinculada', 'predecessor'].filter((key) => (
      Object.values(accScheduleEdits).some((changes) => changes[key] !== undefined)
    ));
    const missingStructuralDefinitions = structuralChanges.filter((key) => !scheduleFieldDefinitionsByKey[key]?.id);
    const structuralWarning = missingStructuralDefinitions.length
      ? `\n\nAtenção: não encontrei no ACC estes campos estruturais: ${missingStructuralDefinitions.map((key) => scheduleFieldLabels[key] || scheduleFieldAliases[key]?.[0] || key).join(', ')}. Eles não serão gravados.`
      : '';
    if (!silent) {
      const approved = window.confirm(`Salvar no CDE/ACC ${manualChangeCount} issue(s) editado(s) e ${selectedSuggestions.length} campo(s) calculado(s)?${structuralWarning}`);
      if (!approved) return;
    }

    setSavingSchedule(true);
    if (silent) {
      setScheduleAutosaveState('saving');
      setScheduleAutosaveMessage('Salvando automaticamente no CDE/ACC...');
    }
    setError('');
    try {
      await requestJson('/api/auth/keep-alive', { method: 'POST' });
      const updatesByIssue = new Map();
      const issuePatchesByIssue = new Map();
      const skippedInvalidOptionFields = new Set();
      const pushCustomAttribute = (issueId, definition, value) => {
        if (!definition?.id) return;
        const prepared = prepareCustomFieldWriteValue(definition, value);
        if (!prepared.valid) {
          skippedInvalidOptionFields.add(`${definition.name || definition.title || definition.displayName || definition.id}: ${value}`);
          return;
        }
        if (!updatesByIssue.has(issueId)) updatesByIssue.set(issueId, []);
        const nextValue = prepared.value;
        updatesByIssue.get(issueId).push({ attributeDefinitionId: definition.id, value: nextValue });
      };

      Object.entries(accScheduleEdits).forEach(([issueId, changes]) => {
        Object.entries(changes).forEach(([key, value]) => {
          if (key === 'title') {
            if (!issuePatchesByIssue.has(issueId)) issuePatchesByIssue.set(issueId, {});
            issuePatchesByIssue.get(issueId).title = value || 'Issue sem título';
            return;
          }
          if (key === 'assignedTo') {
            if (!issuePatchesByIssue.has(issueId)) issuePatchesByIssue.set(issueId, {});
            issuePatchesByIssue.get(issueId).assignedTo = value || null;
            if (value) issuePatchesByIssue.get(issueId).assignedToType = 'user';
            return;
          }
          if (key === 'eapVinculada' && value && !/^\d+(\.\d+)*$/.test(String(value).trim())) {
            skippedInvalidOptionFields.add(`EAP inválida: ${value}`);
            return;
          }
          const definition = scheduleFieldDefinitionsByKey[key];
          if (!definition?.id) return;
          pushCustomAttribute(issueId, definition, getScheduleSaveValue(key, value));
        });
      });

      selectedSuggestions.forEach((suggestion) => {
        pushCustomAttribute(suggestion.issueId, suggestion.definition, getScheduleSaveValue(suggestion.key, suggestion.value));
      });

      if (!updatesByIssue.size && !issuePatchesByIssue.size) {
        if (skippedInvalidOptionFields.size) {
          throw new Error(`Nenhuma alteracao foi enviada porque estes valores nao existem nas listas deste projeto do ACC: ${Array.from(skippedInvalidOptionFields).join(', ')}.`);
        }
        throw new Error('Nao encontrei campos personalizados correspondentes para salvar no ACC. Verifique se os campos usados pelo Cronograma existem neste projeto.');
      }

      const skippedUnmappedFields = new Set();
      let savedScheduleAttributeCount = 0;
      const expectedScheduleAttributeCount = Array.from(updatesByIssue.values()).reduce((total, customAttributes) => total + customAttributes.length, 0);
      const getAttributeLabel = (attributeDefinitionId) => {
        const definition = customFieldDefinitions.find((field) => field.id === attributeDefinitionId)
          || Object.values(scheduleFieldDefinitionsByKey).find((field) => field?.id === attributeDefinitionId);
        return definition?.name || definition?.title || definition?.displayName || attributeDefinitionId;
      };
      const isUnmappedCustomAttributeError = (error) => (
        /deleted or unmapped|unmapped|attribute definition is deleted/i.test(error?.message || '')
      );
      const isInvalidListAttributeError = (error) => (
        /not a valid value for list attribute|invalid value for list|lista|menu suspenso/i.test(error?.message || '')
      );
      const patchScheduleAttributes = async (issueId, customAttributes) => requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues/${encodeURIComponent(issueId)}`, {
        method: 'PATCH',
        body: JSON.stringify({ customAttributes })
      });

      for (const [issueId, customAttributes] of updatesByIssue.entries()) {
        let updatedIssue = null;
        try {
          updatedIssue = await patchScheduleAttributes(issueId, customAttributes);
          savedScheduleAttributeCount += customAttributes.length;
        } catch (batchError) {
          if (!isUnmappedCustomAttributeError(batchError) && !isInvalidListAttributeError(batchError)) throw batchError;
          for (const attribute of customAttributes) {
            try {
              updatedIssue = await patchScheduleAttributes(issueId, [attribute]);
              savedScheduleAttributeCount += 1;
            } catch (fieldError) {
              if (isInvalidListAttributeError(fieldError)) {
                skippedInvalidOptionFields.add(`${getAttributeLabel(attribute.attributeDefinitionId)}: ${attribute.value}`);
                continue;
              }
              if (!isUnmappedCustomAttributeError(fieldError)) throw fieldError;
              skippedUnmappedFields.add(getAttributeLabel(attribute.attributeDefinitionId));
            }
          }
        }
        if (updatedIssue) {
          setIssues((current) => current.map((issue) => (issue.id === issueId ? { ...issue, ...updatedIssue } : issue)));
        }
      }

      if (!savedScheduleAttributeCount && expectedScheduleAttributeCount > 0) {
        throw new Error('O ACC recusou todos os campos enviados porque eles estao apagados ou nao mapeados neste projeto.');
      }

      for (const [issueId, patch] of issuePatchesByIssue.entries()) {
        if (!Object.keys(patch).length) continue;
        const updatedIssue = await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues/${encodeURIComponent(issueId)}`, {
          method: 'PATCH',
          body: JSON.stringify(patch)
        });
        setIssues((current) => current.map((issue) => (issue.id === issueId ? { ...issue, ...updatedIssue } : issue)));
      }

      setScheduleEdits({});
      if (!silent) setScheduleFormulaSelections({});
      await refreshCurrentModule();
      setActiveModule('schedule');
      if (silent) {
        setScheduleAutosaveState('saved');
        setScheduleAutosaveMessage('Alterações salvas automaticamente.');
      }
      if (skippedUnmappedFields.size) {
        setError(`Algumas alteracoes foram salvas, mas o ACC recusou estes campos por estarem apagados ou nao mapeados neste projeto: ${Array.from(skippedUnmappedFields).join(', ')}.`);
      } else if (skippedInvalidOptionFields.size) {
        setError(`As alteracoes validas foram salvas, mas estes valores nao foram enviados porque nao existem nas listas deste projeto do ACC: ${Array.from(skippedInvalidOptionFields).join(', ')}.`);
      }
    } catch (requestError) {
      setError(requestError.message);
      if (silent) {
        setScheduleAutosaveState('error');
        setScheduleAutosaveMessage(requestError.message);
      }
    } finally {
      setSavingSchedule(false);
    }
  }

  function getScheduleColumnLabel(columnKey) {
    const labels = {
      eapAuto: 'EAP',
      typeLabel: 'Tipo do Issue',
      title: 'Título',
      dataContratual: 'Data Contratual',
      dataLimiteInterna: 'Data Limite Interna',
      dataMetaInterna: 'Data Meta Interna',
      dataPublicacao: 'Data Publicacao',
      inicioPlanejado: 'Inicio Planejado',
      terminoPlanejado: 'Término Planejado',
      diasPrevistosAtividade: 'Dias previstos',
      inicioReal: 'Inicio Real',
      terminoReal: 'Termino Real',
      dependencia: 'Dependencia',
      progress: 'Avanço calculado',
      delay: 'Dias de atraso',
      vinculoDependencia: 'Vínculo',
      assignedTo: 'Atribuído',
      eapVinculada: 'EAP Vinculada',
      codigoMarco: 'Código do Marco',
      marcoContratual: 'Marco Contratual',
      codigoDocumentoCliente: 'Código do Documento Cliente',
      codigoDocumentoInterno: 'Código do Documento Interno',
      disciplinaEnvolvida: 'Disciplina envolvida',
      areaResponsavel: 'Área responsável',
      fase: 'Fase',
      tipoItemCronograma: 'Tipo de Item',
      faseFluxo: 'Fase do Fluxo',
      statusEntrega: 'Status da entrega',
      statusCliente: 'Status Cliente',
      statusAnaliseCliente: 'Status Cliente',
      numeroTramitacao: 'Nº da Tramitação',
      dataRealEmissao: 'Data Real de Emissão',
      dataLimiteRetornoCliente: 'Limite Retorno Cliente',
      dataRealRetornoCliente: 'Retorno Cliente',
      dataLimiteRevisaoInterna: 'Limite Revisão Interna',
      dataRealReemissao: 'Reemissão',
      dataAprovacaoFinal: 'Aprovação Final',
      predecessor: 'Predecessor',
      impactoMarco: 'Impacto no Marco',
      impactoCronograma: 'Impacto no Cronograma',
      prioridadeGestao: 'Prioridade de Gestão',
      percentualTecnico: 'Percentual Técnico',
      acaoNecessaria: 'Ação Necessária'
    };
    return labels[columnKey] || scheduleFieldLabels[columnKey] || scheduleFieldAliases[columnKey]?.[0] || columnKey;
  }

  function getScheduleColumnClassName(columnKey) {
    const stickyColumns = ['eapAuto', 'codigoMarco', 'title'];
    return [
      'schedule-col',
      `schedule-col-${columnKey}`,
      stickyColumns.includes(columnKey) ? 'schedule-sticky-col' : ''
    ].filter(Boolean).join(' ');
  }

  function getScheduleColumnHeaderParts(columnKey) {
    const headerParts = {
      codigoMarco: ['Código do', 'Marco'],
      inicioPlanejado: ['Início', 'Planejado'],
      diasPrevistosAtividade: ['Dias', 'previstos'],
      terminoPlanejado: ['Término', 'Planejado'],
      inicioReal: ['Início', 'Real'],
      terminoReal: ['Término', 'Real'],
      dataContratual: ['Data', 'contratual'],
      dataLimiteInterna: ['Data limite', 'interna'],
      statusEntrega: ['Status da', 'entrega'],
      assignedTo: ['Atribuído'],
      vinculoDependencia: ['Vínculo'],
      statusCliente: ['Status', 'Cliente'],
      statusAnaliseCliente: ['Status', 'Cliente'],
      dataPublicacao: ['Data', 'Publicação'],
      acaoNecessaria: ['Ação', 'Necessária'],
      prioridadeGestao: ['Prioridade', 'Gestão'],
      progress: ['Avanço', 'calculado'],
      numeroTramitacao: ['Nº da', 'Tramitação'],
      dataRealEmissao: ['Data real', 'emissão'],
      dataLimiteRetornoCliente: ['Limite retorno', 'cliente'],
      dataRealRetornoCliente: ['Retorno', 'cliente'],
      dataLimiteRevisaoInterna: ['Limite revisão', 'interna'],
      dataAprovacaoFinal: ['Aprovação', 'final']
    };
    return headerParts[columnKey] || [getScheduleColumnLabel(columnKey)];
  }

  function handleCronogramaSort(columnKey) {
    setCronogramaSort((currentSort) => ({
      key: columnKey,
      direction: currentSort.key === columnKey && currentSort.direction === 'asc' ? 'desc' : 'asc'
    }));
  }

  function renderCronogramaSortHeader(columnKey, label) {
    const active = cronogramaSort.key === columnKey;
    return (
      <button
        type="button"
        className={`eap-sort-header ${active ? 'is-active' : ''}`}
        onClick={() => handleCronogramaSort(columnKey)}
        title={`Organizar por ${label}`}
      >
        <span>{label}</span>
        <i aria-hidden="true">{active ? (cronogramaSort.direction === 'asc' ? '↑' : '↓') : '↕'}</i>
      </button>
    );
  }

  function getScheduleRowOwner(row) {
    const issue = row?.issue || {};
    const assignedTo = issue.assignedTo || issue.assignee || issue.assignedToUser || issue.owner || null;
    if (typeof assignedTo === 'string') return assignedTo;
    return assignedTo?.name
      || assignedTo?.displayName
      || assignedTo?.email
      || issue.assignedToName
      || issue.assigneeName
      || row?.fields?.areaResponsavel
      || '';
  }

  function getScheduleRowTypeIcon(row) {
    const type = normalizeText(row?.typeLabel || row?.issue?.issueType || row?.issue?.type || row?.issue?.title);
    if (row?.isMarco || type.includes('marco')) return 'M';
    if (type.includes('risco') || type.includes('restricao')) return '!';
    if (type.includes('informacao') || type.includes('solicitacao')) return '?';
    if (type.includes('emissao') || type.includes('cliente')) return 'E';
    return 'T';
  }

  function resolveSchedulePredecessorDisplayValue(row) {
    const rawValue = normalizePredecessorInput(scheduleEdits[row.id]?.predecessor ?? row.fields.predecessor);
    const predecessorIds = parsePredecessorIds(rawValue);
    if (!predecessorIds.length) return rawValue;

    return predecessorIds.map((predecessorId) => {
      const normalizedPredecessorId = normalizeText(predecessorId);
      const matchedPredecessor =
        row.predecessorIssues?.find((issue) =>
          String(issue.id) === String(predecessorId)
          || normalizeText(getIssueHumanId(issue)) === normalizedPredecessorId
        )
        || schedulePlannerRows.find((candidate) =>
          String(candidate.issue?.id) === String(predecessorId)
          || normalizeText(getIssueHumanId(candidate.issue)) === normalizedPredecessorId
        )?.issue;

      return matchedPredecessor ? getIssueHumanId(matchedPredecessor) : predecessorId;
    }).join(', ');
  }

  function getScheduleCellValue(row, columnKey) {
    const edited = scheduleEdits[row.id] || {};
    const dependencyEnabled = normalizeText(edited.dependencia || row.localDependency || '').startsWith('s');
    const predecessorIssue = row.predecessorIssues?.[0];
    if (dependencyEnabled && predecessorIssue && columnKey === 'inicioPlanejado') {
      const predecessorStart = getScheduleField(predecessorIssue, customFieldDefinitions, 'inicioPlanejado');
      const predecessorDays = getScheduleField(predecessorIssue, customFieldDefinitions, 'diasPrevistosAtividade');
      const predecessorEnd = getScheduleField(predecessorIssue, customFieldDefinitions, 'terminoPlanejado')
        || computeScheduleEndDateFromBusinessDays(predecessorStart, predecessorDays)
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataMetaInterna')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataContratual');
      const nextStart = parseScheduleDate(predecessorEnd) ? formatBrazilianDate(addBusinessDays(parseScheduleDate(predecessorEnd), 1)) : '';
      if (nextStart) return nextStart;
    }
    if (dependencyEnabled && predecessorIssue && columnKey === 'inicioReal') {
      const predecessorRealEnd = getScheduleField(predecessorIssue, customFieldDefinitions, 'terminoReal')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataAprovacaoFinal')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataRealReemissao')
        || getScheduleField(predecessorIssue, customFieldDefinitions, 'dataRealRetornoCliente');
      const nextStart = parseScheduleDate(predecessorRealEnd) ? formatBrazilianDate(addBusinessDays(parseScheduleDate(predecessorRealEnd), 1)) : '';
      if (nextStart) return nextStart;
    }
    if (columnKey === 'terminoPlanejado' && edited.terminoPlanejado === undefined) {
      const computedEnd = computeScheduleEndDateFromBusinessDays(
        getScheduleCellValue(row, 'inicioPlanejado'),
        edited.diasPrevistosAtividade ?? row.fields.diasPrevistosAtividade
      );
      return computedEnd || row.fields.terminoPlanejado || '';
    }
    if (columnKey === 'eapAuto') return edited.eapVinculada ?? row.eapAuto ?? row.fields.eapVinculada ?? '';
    if (columnKey === 'title') return edited.title ?? row.issue.title ?? 'Issue sem título';
    if (columnKey === 'typeLabel') return row.typeLabel;
    if (columnKey === 'progress') {
      if (row.isMarco && row.children?.length && edited.statusEntrega === undefined) {
        return `${Math.round(row.progress || 0)}%`;
      }
      return `${Math.round(getScheduleProgressFromDeliveryStatus(edited.statusEntrega ?? row.fields.statusEntrega) || 0)}%`;
    }
    if (columnKey === 'delay') {
      const fields = {
        ...row.fields,
        ...edited,
        inicioPlanejado: getScheduleCellValue(row, 'inicioPlanejado'),
        terminoPlanejado: getScheduleCellValue(row, 'terminoPlanejado')
      };
      return calculateScheduleDelay(fields, isScheduleCompleted(row.issue, customFieldDefinitions, fields)).days || 0;
    }
    if (columnKey === 'dependencia') return edited.dependencia ?? row.localDependency ?? row.fields.dependencia ?? 'Nao';
    if (columnKey === 'predecessor') return resolveSchedulePredecessorDisplayValue(row);
    if (columnKey === 'assignedTo') return edited.assignedTo ?? getKanbanAssignedUserId(row.issue);
    return edited[columnKey] ?? row.fields[columnKey] ?? '';
  }

  function getScheduleFieldOptions(columnKey) {
    const definition = scheduleFieldDefinitionsByKey[columnKey];
    const optionSources = [
      definition?.options,
      definition?.values,
      definition?.allowedValues,
      definition?.choices,
      definition?.enumValues
    ].filter(Array.isArray);
    const options = optionSources.flat().map((option) => {
      if (typeof option === 'string') return { value: option, label: option };
      const label = option.label || option.name || option.title || option.displayName || option.value || option.id;
      const value = option.value || option.label || option.name || option.title || option.displayName || option.id;
      return label && value ? { value: String(value), label: String(label) } : null;
    }).filter(Boolean);
    return Array.from(new Map(options.map((option) => [option.label, option])).values());
  }

  function renderScheduleDetailEditor(row, columnKey, label) {
    if (!row) return null;
    const value = getScheduleCellValue(row, columnKey);
    const isDateColumn = ['dataContratual', 'dataLimiteInterna', 'dataMetaInterna', 'dataPublicacao', 'inicioPlanejado', 'dataRealEmissao', 'dataLimiteRetornoCliente', 'dataRealRetornoCliente', 'dataLimiteRevisaoInterna', 'dataRealReemissao', 'dataAprovacaoFinal', 'inicioReal', 'terminoReal'].includes(columnKey);
    const isNumberColumn = columnKey === 'diasPrevistosAtividade' || columnKey === 'delay';
    const options = ['statusEntrega', 'prioridadeGestao', 'faseFluxo', 'statusCliente', 'disciplinaEnvolvida'].includes(columnKey)
      ? getScheduleFieldOptions(columnKey)
      : [];
    const disabled = columnKey === 'delay' || columnKey === 'terminoPlanejado' || (columnKey === 'statusEntrega' && !options.length);
    return (
      <label className="schedule-detail-field">
        <span>{label}</span>
        {options.length || columnKey === 'statusEntrega' ? (
          <select
            value={value}
            disabled={disabled}
            onChange={(event) => editScheduleCell(row.id, columnKey, event.target.value)}
          >
            <option value="">Preencher</option>
            {value && !options.some((option) => option.value === value || option.label === value) && <option value={value}>{value}</option>}
            {options.map((option) => <option key={`${columnKey}:${option.value}`} value={option.value}>{option.label}</option>)}
          </select>
        ) : (
          <input
            type={isDateColumn ? 'date' : isNumberColumn ? 'number' : 'text'}
            min={isNumberColumn ? '0' : undefined}
            step={isNumberColumn ? '1' : undefined}
            value={isDateColumn ? getScheduleDateInputValue(value) : value}
            disabled={disabled}
            onChange={(event) => editScheduleCell(row.id, columnKey, event.target.value)}
            placeholder="Preencher"
          />
        )}
      </label>
    );
  }

  function renderScheduleAssigneeEditor(row) {
    if (!row) return null;
    const value = getScheduleCellValue(row, 'assignedTo');
    const currentLabel = getScheduleRowOwner(row);
    const hasCurrentUser = projectUsers.some((projectUser) => String(projectUser.id) === String(value));
    return (
      <label className="schedule-detail-field">
        <span>Atribuido</span>
        <select
          value={value}
          disabled={!row?.issue?.id || savingIssueId === row.issue.id}
          onChange={(event) => updateScheduleIssueNative(row, event.target.value ? { assignedTo: event.target.value, assignedToType: 'user' } : { assignedTo: null })}
        >
          <option value="">Sem atribuicao</option>
          {value && !hasCurrentUser && <option value={value}>{currentLabel || value}</option>}
          {projectUsers.map((projectUser) => (
            <option key={projectUser.id} value={projectUser.id}>
              {projectUser.name}{projectUser.email ? ` - ${projectUser.email}` : ''}
            </option>
          ))}
        </select>
      </label>
    );
  }

  function getScheduleIssueCategoryId(row) {
    const issue = row?.issue || {};
    const raw = issue.raw || {};
    const currentId = issue.issueTypeId || raw.issueTypeId || raw.typeId || raw.attributes?.issueTypeId || raw.attributes?.typeId;
    if (currentId) return currentId;
    const label = issue.category || issue.issueType || row?.category || row?.typeLabel;
    return issueCategories.find((item) => normalizeText(item.title) === normalizeText(label) || normalizeText(item.category) === normalizeText(label))?.id || '';
  }

  function getScheduleIssueSubtypeId(row, categoryId) {
    const issue = row?.issue || {};
    const raw = issue.raw || {};
    const currentId = issue.issueSubtypeId || raw.issueSubtypeId || raw.subtypeId || raw.attributes?.issueSubtypeId || raw.attributes?.subtypeId;
    if (currentId) return currentId;
    const label = issue.issueType || issue.issueSubtype || row?.typeLabel;
    return issueTypeOptions.find((item) => item.kind === 'subtype' && (!categoryId || item.typeId === categoryId) && normalizeText(item.title) === normalizeText(label))?.id || '';
  }

  async function updateScheduleIssueNative(row, patch) {
    if (!row?.issue?.id) {
      setError('Selecione uma issue valida antes de atualizar o ACC.');
      return;
    }
    await updateIssue(row.issue.id, patch);
  }

  function renderScheduleNativeStatusEditor(row) {
    const value = row?.issue?.status || '';
    return (
      <label className="schedule-detail-field">
        <span>Status do Issue</span>
        <select
          value={value}
          disabled={!row?.issue?.id || savingIssueId === row.issue.id}
          onChange={(event) => updateScheduleIssueNative(row, { status: event.target.value })}
        >
          <option value="">Preencher</option>
          {value && !statusOptions.some((option) => option.value === value) && <option value={value}>{value}</option>}
          {statusOptions.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
        </select>
      </label>
    );
  }

  function renderScheduleIssueCategoryEditor(row) {
    const value = getScheduleIssueCategoryId(row);
    const currentLabel = row?.issue?.category || row?.typeLabel;
    return (
      <label className="schedule-detail-field">
        <span>Categoria</span>
        <select
          value={value}
          disabled={!row?.issue?.id || savingIssueId === row.issue.id || !issueCategories.length}
          onChange={(event) => updateScheduleIssueNative(row, { issueTypeId: event.target.value, issueSubtypeId: '' })}
        >
          <option value="">{currentLabel && !value ? currentLabel : 'Preencher'}</option>
          {issueCategories.map((option) => <option key={option.id} value={option.id}>{option.title || option.category}</option>)}
        </select>
      </label>
    );
  }

  function renderScheduleIssueTypeEditor(row) {
    const categoryId = getScheduleIssueCategoryId(row);
    const value = getScheduleIssueSubtypeId(row, categoryId);
    const options = issueTypeOptions.filter((item) => item.kind === 'subtype' && (!categoryId || item.typeId === categoryId));
    const currentLabel = row?.issue?.issueType || row?.typeLabel;
    return (
      <label className="schedule-detail-field">
        <span>Tipo do Issue</span>
        <select
          value={value}
          disabled={!row?.issue?.id || savingIssueId === row.issue.id || !options.length}
          onChange={(event) => updateScheduleIssueNative(row, { issueSubtypeId: event.target.value })}
        >
          <option value="">{currentLabel && !value ? currentLabel : 'Preencher'}</option>
          {options.map((option) => <option key={option.id} value={option.id}>{option.title}</option>)}
        </select>
      </label>
    );
  }

  function downloadScheduleCsv(fileName, headers, rows) {
    const csv = [headers, ...rows].map((row) => row.map(csvEscape).join(';')).join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = fileName;
    link.click();
    window.URL.revokeObjectURL(url);
  }

  function exportSchedulePowerBiCsv() {
    const rows = schedulePlannerRows;
    if (!rows.length) {
      setError('Nao ha itens de cronograma para exportar.');
      return;
    }
    const headers = scheduleColumns.map((column) => getScheduleColumnLabel(column));
    const csvRows = rows.map((row) => scheduleColumns.map((column) => getScheduleCellValue(row, column)));
    downloadScheduleCsv(`central-g5-cronograma-powerbi-${selectedProject?.name || selectedProjectId || 'projeto'}.csv`, headers, csvRows);
  }

  function exportScheduleMsProjectCsv() {
    const rows = schedulePlannerRows;
    if (!rows.length) {
      setError('Nao ha itens de cronograma para exportar.');
      return;
    }
    const headers = ['ID', 'Nome da tarefa', 'Inicio', 'Termino', 'Duracao', 'Predecessoras', 'Percentual concluido', 'Nivel da estrutura', 'Observacoes'];
    const csvRows = rows.map((row) => [
      row.eapAuto || getScheduleCellValue(row, 'eapAuto'),
      getScheduleCellValue(row, 'title'),
      getScheduleCellValue(row, 'inicioPlanejado') || getScheduleCellValue(row, 'dataContratual'),
      getScheduleCellValue(row, 'terminoPlanejado') || getScheduleCellValue(row, 'dataMetaInterna'),
      getScheduleCellValue(row, 'diasPrevistosAtividade'),
      getScheduleCellValue(row, 'predecessor'),
      `${Math.round(row.progress || 0)}%`,
      String((row.plannerLevel || 0) + 1),
      getScheduleCellValue(row, 'acaoNecessaria')
    ]);
    downloadScheduleCsv(`central-g5-cronograma-msproject-${selectedProject?.name || selectedProjectId || 'projeto'}.csv`, headers, csvRows);
  }

  function renderScheduleDependencyBadge(row) {
    const state = getScheduleDependencyState(row);
    return (
      <span className={`schedule-dependency-badge is-${state.status}`} title={state.detail || state.label}>
        <strong>{state.label}</strong>
        {state.detail ? <small>{state.detail}</small> : null}
      </span>
    );
  }

  function renderScheduleAssigneeCell(row) {
    const value = getScheduleCellValue(row, 'assignedTo');
    const currentLabel = getScheduleRowOwner(row);
    const hasCurrentUser = projectUsers.some((projectUser) => String(projectUser.id) === String(value));
    return (
      <select
        className="cronograma-cell-input schedule-cell-input schedule-clean-input"
        value={value}
        onChange={(event) => editScheduleCell(row.id, 'assignedTo', event.target.value)}
      >
        <option value="">Sem atribuição</option>
        {value && !hasCurrentUser && <option value={value}>{currentLabel || value}</option>}
        {projectUsers.map((projectUser) => (
          <option key={projectUser.id} value={projectUser.id}>
            {projectUser.name}{projectUser.email ? ` - ${projectUser.email}` : ''}
          </option>
        ))}
      </select>
    );
  }

  function renderScheduleCell(row, columnKey) {
    if (columnKey === 'progress') {
      const progressValue = Number(String(getScheduleCellValue(row, columnKey)).replace(/[^\d.-]/g, '')) || 0;
      return (
        <div className="execution-cell" title="Calculado automaticamente pelo campo Status da entrega">
          <span>{Math.round(progressValue)}%</span>
          <div className="execution-bar"><i style={{ width: `${Math.round(progressValue)}%` }} /></div>
        </div>
      );
    }
    if (columnKey === 'delay') {
      const delayDays = Number(getScheduleCellValue(row, columnKey)) || 0;
      return <span className={delayDays > 0 ? 'schedule-delay-alert' : ''}>{delayDays}</span>;
    }
    if (columnKey === 'terminoPlanejado') {
      return <span className="schedule-calculated-cell">{getScheduleCellValue(row, columnKey) || '-'}</span>;
    }
    if (columnKey === 'eapAuto') {
      if (scheduleViewMode === 'coordenacao') {
        return (
          <input
            className="cronograma-cell-input schedule-cell-input schedule-clean-input schedule-eap-editor"
            value={getScheduleCellValue(row, columnKey)}
            onChange={(event) => editScheduleCell(row.id, 'eapVinculada', event.target.value)}
            placeholder="EAP"
          />
        );
      }
      return <span className="schedule-eap-code">{getScheduleCellValue(row, columnKey) || '-'}</span>;
    }
    if (columnKey === 'title') {
      if (scheduleViewMode === 'coordenacao') {
        return (
          <span className="schedule-title-cell" style={{ paddingLeft: `${Math.max(0, row.plannerLevel || 0) * 18}px` }}>
            {row.plannerLevel ? <i aria-hidden="true">↳</i> : null}
            <em aria-hidden="true" className="schedule-type-icon">{getScheduleRowTypeIcon(row)}</em>
            <input
              className="cronograma-cell-input schedule-cell-input schedule-clean-input schedule-title-editor"
              value={getScheduleCellValue(row, columnKey)}
              onChange={(event) => editScheduleCell(row.id, 'title', event.target.value)}
              placeholder="Título"
            />
          </span>
        );
      }
      return (
        <span className="schedule-title-cell" style={{ paddingLeft: `${Math.max(0, row.plannerLevel || 0) * 18}px` }}>
          {row.plannerLevel ? <i aria-hidden="true">↳</i> : null}
          <em aria-hidden="true" className="schedule-type-icon">{getScheduleRowTypeIcon(row)}</em>
          <strong>{getScheduleCellValue(row, columnKey)}</strong>
        </span>
      );
    }
    if (columnKey === 'assignedTo') {
      return renderScheduleAssigneeCell(row);
    }
    if (columnKey === 'vinculoDependencia') {
      return renderScheduleDependencyBadge(row);
    }
    if (columnKey === 'dependencia') {
      return (
        <select
          className="cronograma-cell-input schedule-cell-input"
          value={getScheduleCellValue(row, columnKey)}
          onChange={(event) => editScheduleDependency(row, event.target.value)}
        >
          <option value="Nao">Nao</option>
          <option value="Sim">Sim</option>
        </select>
      );
    }
    if (scheduleEditableKeys.includes(columnKey)) {
      const value = getScheduleCellValue(row, columnKey);
      const isDateColumn = ['dataContratual', 'dataLimiteInterna', 'dataMetaInterna', 'dataPublicacao', 'inicioPlanejado', 'dataRealEmissao', 'dataLimiteRetornoCliente', 'dataRealRetornoCliente', 'dataLimiteRevisaoInterna', 'dataRealReemissao', 'dataAprovacaoFinal', 'inicioReal', 'terminoReal'].includes(columnKey);
      const isNumberColumn = columnKey === 'diasPrevistosAtividade';
      const options = ['statusEntrega', 'prioridadeGestao', 'faseFluxo', 'statusCliente'].includes(columnKey)
        ? getScheduleFieldOptions(columnKey)
        : [];
      if (scheduleViewMode === 'coordenacao' && columnKey === 'predecessor') {
        const state = getScheduleDependencyState(row);
        return (
          <span className="schedule-predecessor-cell">
            <input
              className="cronograma-cell-input schedule-cell-input schedule-clean-input"
              type="text"
              value={value}
              onChange={(event) => editScheduleCell(row.id, columnKey, event.target.value)}
              placeholder="ID"
            />
            {value ? <small className={`schedule-predecessor-hint is-${state.status}`}>{state.detail || state.label}</small> : null}
          </span>
        );
      }
      if (options.length || columnKey === 'statusEntrega') {
        const hasCurrentValue = value && !options.some((option) => option.value === value || option.label === value);
        return (
          <select
            className={`cronograma-cell-input schedule-cell-input ${scheduleViewMode === 'coordenacao' ? 'schedule-clean-input' : ''}`}
            value={value}
            disabled={columnKey === 'statusEntrega' && !options.length}
            title={columnKey === 'statusEntrega' && !options.length ? 'As opções deste campo ainda não foram carregadas do ACC.' : undefined}
            onChange={(event) => editScheduleCell(row.id, columnKey, event.target.value)}
          >
            <option value="">Preencher</option>
            {hasCurrentValue && <option value={value}>{value}</option>}
            {options.map((option) => (
              <option key={`${columnKey}:${option.value}`} value={option.value}>{option.label}</option>
            ))}
          </select>
        );
      }
      return (
        <input
          className={`cronograma-cell-input schedule-cell-input ${scheduleViewMode === 'coordenacao' ? 'schedule-clean-input' : ''}`}
          type={isDateColumn ? 'date' : isNumberColumn ? 'number' : 'text'}
          min={isNumberColumn ? '0' : undefined}
          step={isNumberColumn ? '1' : undefined}
          value={isDateColumn ? getScheduleDateInputValue(value) : value}
          onChange={(event) => editScheduleCell(row.id, columnKey, event.target.value)}
          placeholder="Preencher"
        />
      );
    }
    return <span>{getScheduleCellValue(row, columnKey) || '-'}</span>;
  }

  async function loadIssueReportFromAcc(force = false) {
    if (!selectedProjectId) return;
    setIssueReportLoading(true);
    setIssueReportStatus('');
    try {
      const cacheKey = `cronograma-issue-report:${selectedProjectId}`;
      if (!force) {
        const cached = window.localStorage.getItem(cacheKey);
        if (cached) {
          const parsed = JSON.parse(cached);
          setIssueReportInfo(parsed);
          setIssueReportStatus('Relatório carregado do cache local.');
          setIssueReportLoading(false);
          return;
        }
      }
      const payload = await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/cronograma/issue-report`);
      setIssueReportInfo(payload.file);
      window.localStorage.setItem(cacheKey, JSON.stringify(payload.file));
      setIssueReportStatus('Relatório do ACC localizado com sucesso.');
    } catch (requestError) {
      setIssueReportInfo(null);
      setIssueReportStatus('Relatório de issues não encontrado na pasta configurada. O cronograma será montado somente com dados da API.');
      console.error('[Modulo 8][Cronograma] Erro ao localizar relatório de issues.', requestError);
    } finally {
      setIssueReportLoading(false);
    }
  }

  function getCronogramaExportRows() {
    return (cronogramaRowsFiltered.length ? cronogramaRowsFiltered : cronogramaRows)
      .filter((row) => row?.issue)
      .sort((firstRow, secondRow) =>
        String(firstRow.eapVinculada || '').localeCompare(String(secondRow.eapVinculada || ''), 'pt-BR', { numeric: true, sensitivity: 'base' })
        || String(firstRow.titulo || '').localeCompare(String(secondRow.titulo || ''), 'pt-BR', { sensitivity: 'base' })
      );
  }

  function xmlEscape(value) {
    return String(value ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }

  function getCronogramaExportDate(row, preferredKey) {
    const editedRow = cronogramaEdits[row?.issue?.id] || {};
    if (preferredKey === 'terminoPrevisto') {
      const computedFinish = computeEndDateFromStartAndDays(
        editedRow.inicioPrevisto ?? row?.inicioPrevisto,
        editedRow.tempoAtividade ?? row?.tempoAtividade
      );
      if (computedFinish) return parseIssueDate(computedFinish);
    }
    const preferred = parseIssueDate(editedRow?.[preferredKey] ?? row?.[preferredKey]);
    const fallback = parseIssueDate(editedRow?.terminoPrevisto ?? row?.terminoPrevisto) || parseIssueDate(editedRow?.inicioPrevisto ?? row?.inicioPrevisto) || parseIssueDate(row?.dataPrevistaG5) || parseIssueDate(row?.dataContratual) || parseIssueDate(row?.startDate) || new Date();
    return preferred || fallback;
  }

  function formatProjectDate(dateValue) {
    const date = parseIssueDate(dateValue) || dateValue || new Date();
    const safeDate = date instanceof Date ? date : new Date(date);
    if (Number.isNaN(safeDate.getTime())) return `${new Date().toISOString().slice(0, 10)}T08:00:00`;
    return `${safeDate.toISOString().slice(0, 10)}T08:00:00`;
  }

  function getCronogramaDuration(row) {
    const editedValue = cronogramaEdits[row.issue.id]?.tempoAtividade;
    const rawValue = editedValue ?? row.tempoAtividade;
    const numericValue = Number(String(rawValue || '').replace(',', '.').replace(/[^\d.-]/g, ''));
    const days = Number.isFinite(numericValue) && numericValue > 0 ? Math.max(1, Math.round(numericValue)) : 1;
    return `PT${days * 8}H0M0S`;
  }

  function getCronogramaDurationDays(row) {
    const editedValue = cronogramaEdits[row.issue.id]?.tempoAtividade;
    const rawValue = editedValue ?? row.tempoAtividade;
    const numericValue = Number(String(rawValue || '').replace(',', '.').replace(/[^\d.-]/g, ''));
    return Number.isFinite(numericValue) && numericValue > 0 ? Math.max(1, Math.round(numericValue)) : 1;
  }

  function formatMsProjectCsvDate(dateValue) {
    const date = parseIssueDate(dateValue);
    if (!date || Number.isNaN(date.getTime())) return '';
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    return `${day}/${month}/${date.getFullYear()}`;
  }

  function exportCronogramaMsProjectCsv() {
    const rows = getCronogramaExportRows();
    if (!rows.length) {
      setError('Nao ha itens de cronograma para exportar.');
      return;
    }

    const taskIdByKey = new Map();
    rows.forEach((row, index) => {
      const uid = index + 1;
      [
        row.issue.id,
        row.issue.displayId,
        row.issue.autodeskId,
        row.eapVinculada,
        row.titulo
      ].filter(Boolean).forEach((key) => taskIdByKey.set(String(key).trim(), uid));
    });

    const headers = [
      'ID',
      'Task Name',
      'WBS',
      'Outline Level',
      'Start',
      'Finish',
      'Duration',
      'Percent Complete',
      'Predecessors',
      'Resource Names',
      'Notes',
      'Text1 ACC ID',
      'Text2 Marco',
      'Text3 Status ACC',
      'Text4 Categoria',
      'Text5 Tipo'
    ];

    const csvRows = rows.map((row, index) => {
      const uid = index + 1;
      const start = getCronogramaExportDate(row, 'inicioPrevisto');
      const finish = getCronogramaExportDate(row, 'terminoPrevisto');
      const outlineLevel = Math.max(1, String(row.eapVinculada || '').split('.').filter(Boolean).length);
      const predecessors = (row.predecessorIds || [])
        .map((predecessorId) => taskIdByKey.get(String(predecessorId).trim()))
        .filter((predecessorUid) => predecessorUid && predecessorUid !== uid)
        .join(',');

      return [
        uid,
        `${row.eapVinculada || uid} - ${row.titulo || 'Sem titulo'}`,
        row.eapVinculada || uid,
        outlineLevel,
        formatMsProjectCsvDate(start),
        formatMsProjectCsvDate(finish),
        `${getCronogramaDurationDays(row)}d`,
        Math.max(0, Math.min(100, Number(row.executado || 0))),
        predecessors,
        row.atribuidoA || '',
        [
          `Marco: ${row.marcoContratual || '-'}`,
          `Status: ${row.statusValue || row.issue.status || '-'}`,
          `Atribuido a: ${row.atribuidoA || '-'}`,
          `ACC ID: ${row.issue.displayId || row.issue.id || '-'}`
        ].join('\n'),
        row.issue.displayId || row.issue.id || '',
        row.marcoContratual || '',
        row.statusValue || row.issue.status || '',
        row.issue.category || '',
        row.issue.issueType || row.issue.type || row.issue.issueSubtype || ''
      ];
    });

    const csv = [headers, ...csvRows].map((row) => row.map(csvEscape).join(';')).join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `central-g5-eap-msproject-${selectedProject?.name || selectedProjectId || 'projeto'}.csv`;
    link.click();
    window.URL.revokeObjectURL(url);
  }

  function exportCronogramaPowerBiCsv() {
    const rows = getCronogramaExportRows();
    if (!rows.length) {
      setError('Nao ha itens de cronograma para exportar.');
      return;
    }

    const headers = [
      'Projeto',
      'EAP',
      'Marco contratual',
      'Titulo',
      'Categoria',
      'Tipo',
      'Status ACC',
      'Status gerencial',
      'Inicio Planejado',
      'Término Planejado',
      'Inicio real',
      'Termino real',
      'Data contratual',
      'Atribuido a',
      'Predecessores',
      'Tempo para atividade',
      'Percentual executado',
      'Concluido',
      'Atrasado',
      'Bloqueado',
      'Dias de atraso',
      'Issue ID'
    ];

    const csvRows = rows.map((row) => [
      selectedProject?.name || selectedProjectId || '',
      row.eapVinculada || '',
      row.marcoContratual || '',
      row.titulo || '',
      row.issue.category || '',
      row.issue.issueType || row.issue.type || row.issue.issueSubtype || '',
      row.statusValue || row.issue.status || '',
      row.statusConsolidado || '',
      getDateInputValue(row.inicioPrevisto) || '',
      getDateInputValue(row.terminoPrevisto) || '',
      getDateInputValue(row.inicioReal) || '',
      getDateInputValue(row.terminoReal) || '',
      getDateInputValue(row.dataContratual) || '',
      row.atribuidoA || '',
      row.predecessoresTexto || '',
      cronogramaEdits[row.issue.id]?.tempoAtividade ?? row.tempoAtividade ?? '',
      Number(row.executado || 0),
      row.concluded ? 'Sim' : 'Nao',
      row.overdue ? 'Sim' : 'Nao',
      row.blocked ? 'Sim' : 'Nao',
      row.delayDays ?? '',
      row.issue.displayId || row.issue.id || ''
    ]);

    const csv = [headers, ...csvRows].map((row) => row.map(csvEscape).join(';')).join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `central-g5-powerbi-eap-${selectedProject?.name || selectedProjectId || 'projeto'}.csv`;
    link.click();
    window.URL.revokeObjectURL(url);
  }

  useEffect(() => {
    if (!issues.length) return;
    const relationships = Object.values(issueRelationshipsById || {});
    const referencesFound = relationships.reduce((total, item) => total + (Array.isArray(item) ? item.length : 0), 0);
    const withoutMarco = issues.filter((issue) => !getCustomFieldValue(issue, customFieldDefinitions, 'Código do Marco')).length;
    const withMarco = issues.length - withoutMarco;
    const overdue = issues.filter(isOverdue).length;
    const blocked = issues.filter((issue) => {
      const predecessors = resolveIssueReferences(issue, buildIssueLookupIndexes(issues));
      return predecessors.length > 0 && predecessors.some((predecessor) => isOpenIssue(predecessor));
    }).length;
    console.info('[Modulo 3] Diagnostico ACC', {
      totalIssues: issues.length,
      customFieldDefinitions: customFieldDefinitions.length,
      withMarco,
      withoutMarco,
      referencesFound,
      issuesWithReferences: issues.filter((issue) => (issue.relationshipReferenceIds || []).length > 0).length,
      issuesWithoutReferences: issues.filter((issue) => (issue.relationshipReferenceIds || []).length === 0).length,
      overdue,
      blocked
    });
  }, [issues, issueRelationshipsById, customFieldDefinitions]);

  const monthlyTimeline = useMemo(() => {
    const groupedIssues = new Map();

    for (const issue of visibleIssues) {
      const monthKey = getMonthKey(issue.dueDate);
      const monthIssues = groupedIssues.get(monthKey) || [];
      monthIssues.push(issue);
      groupedIssues.set(monthKey, monthIssues);
    }

    return [...groupedIssues.entries()]
      .sort(([firstMonth], [secondMonth]) => {
        if (firstMonth === 'sem-prazo') return 1;
        if (secondMonth === 'sem-prazo') return -1;
        return String(firstMonth).localeCompare(String(secondMonth));
      })
      .map(([monthKey, monthIssues]) => ({
        monthKey,
        label: getMonthLabel(monthKey),
        total: monthIssues.length,
        open: monthIssues.filter(isOpenIssue).length,
        overdue: monthIssues.filter(isOverdue).length,
        issues: [...monthIssues].sort(sortByName)
      }));
  }, [visibleIssues]);

  const groupedIssues = useMemo(() => {
    const groups = new Map();

    const issuesForModule =
      activeModule === 'interface'
        ? visibleIssues.filter((issue) =>
            normalizeText([issue.category, issue.issueType, issue.title].filter(Boolean).join(' ')).includes(
              normalizeText('Interface e Coordenação Multidisciplinar')
            )
          )
        : visibleIssues;

    for (const issue of issuesForModule) {
      const groupName = issue.issueType || 'Sem tipo de issue';
      const groupIssues = groups.get(groupName) || [];
      groupIssues.push(issue);
      groups.set(groupName, groupIssues);
    }

    return [...groups.entries()]
      .sort(([firstGroup], [secondGroup]) => String(firstGroup).localeCompare(String(secondGroup), 'pt-BR', { sensitivity: 'base' }))
      .map(([name, groupIssues]) => ({
        name,
        issues: groupIssues
      }));
  }, [visibleIssues, activeModule]);

  const interfacesPendenciasDashboard = useMemo(() => {
    const allIssues = visibleIssues;
    const impactKinds = [
      { key: 'prazo', label: 'Prazo', aliases: ['impacto no cronograma', 'impacto no cronograma:', 'impacto em prazo', 'impacto no prazo', 'impacto previsto'] },
      { key: 'escopo', label: 'Escopo', aliases: ['impacto no escopo', 'impacto em escopo'] },
      { key: 'medicao', label: 'Medição', aliases: ['impacto em medição', 'impacto em medicao', 'impacto na medição', 'impacto na medicao'] },
      { key: 'qualidade', label: 'Qualidade', aliases: ['impacto na qualidade', 'impacto em qualidade', 'qualidade do modelo bim'] }
    ];
    const levels = ['Alto', 'Médio', 'Baixo', 'Sem classificação'];
    const getIssueImpactItems = (issue) => impactKinds.map((impact) => {
      const rawValue = getIssueCustomValue(issue, impact.aliases);
      const level = normalizeImpactLevel(rawValue);
      return { ...impact, rawValue, level };
    });
    const disciplineOptions = [...new Set(allIssues.map((issue) => getIssueDiscipline(issue) || 'Sem disciplina'))]
      .sort((first, second) => String(first).localeCompare(String(second), 'pt-BR', { sensitivity: 'base' }));
    const sourceIssues = allIssues.filter((issue) => {
      const discipline = getIssueDiscipline(issue) || 'Sem disciplina';
      if (interfacesDisciplineFilter !== 'all' && discipline !== interfacesDisciplineFilter) return false;
      if (interfacesStatusFilter === 'open' && !isOpenIssue(issue)) return false;
      if (interfacesStatusFilter === 'closed' && isOpenIssue(issue)) return false;
      if (interfacesStatusFilter === 'overdue' && !isOverdue(issue)) return false;
      if (interfacesImpactFilter !== 'all' && !getIssueImpactItems(issue).some((impact) => impact.level === interfacesImpactFilter)) return false;
      return true;
    });
    const disciplineMap = new Map();
    const impactSummary = impactKinds.map((kind) => ({
      ...kind,
      counts: Object.fromEntries(levels.map((level) => [level, 0]))
    }));

    sourceIssues.forEach((issue) => {
      const discipline = getIssueDiscipline(issue) || 'Sem disciplina';
      const group = disciplineMap.get(discipline) || {
        discipline,
        total: 0,
        open: 0,
        overdue: 0,
        highImpact: 0,
        meetings: 0,
        encaminhamentos: 0,
        issues: []
      };
      const issueText = normalizeText([issue.title, issue.issueType, issue.category, issue.description].filter(Boolean).join(' '));
      const hasMeetingContext = issueText.includes('reuniao') || issueText.includes('ata') || issueText.includes('encaminhamento');
      const hasOwner = Boolean(issue.assignedTo || issue.assignee || issue.assignedUsers || issue.raw?.assignedTo);
      const issueImpacts = getIssueImpactItems(issue);
      issueImpacts.forEach((issueImpact) => {
        const impact = impactSummary.find((item) => item.key === issueImpact.key);
        if (impact) impact.counts[issueImpact.level] = (impact.counts[issueImpact.level] || 0) + 1;
      });

      group.total += 1;
      group.open += isOpenIssue(issue) ? 1 : 0;
      group.overdue += isOverdue(issue) ? 1 : 0;
      group.highImpact += issueImpacts.some((impact) => impact.level === 'Alto') ? 1 : 0;
      group.meetings += hasMeetingContext ? 1 : 0;
      group.encaminhamentos += isOpenIssue(issue) && (hasOwner || hasMeetingContext) ? 1 : 0;
      group.issues.push(issue);
      disciplineMap.set(discipline, group);
    });

    const disciplines = Array.from(disciplineMap.values())
      .sort((first, second) => second.open - first.open || String(first.discipline).localeCompare(String(second.discipline), 'pt-BR', { sensitivity: 'base' }));
    const highlightedIssues = [...sourceIssues]
      .sort((firstIssue, secondIssue) => {
        const firstOverdue = isOverdue(firstIssue) ? 1 : 0;
        const secondOverdue = isOverdue(secondIssue) ? 1 : 0;
        if (firstOverdue !== secondOverdue) return secondOverdue - firstOverdue;
        return getIssueTimelineDate(firstIssue) - getIssueTimelineDate(secondIssue);
      })
      .slice(0, 8);

    return {
      total: sourceIssues.length,
      totalProject: allIssues.length,
      open: sourceIssues.filter(isOpenIssue).length,
      overdue: sourceIssues.filter(isOverdue).length,
      disciplines,
      disciplineOptions,
      impactSummary,
      meetings: disciplines.reduce((sum, discipline) => sum + discipline.meetings, 0),
      encaminhamentos: disciplines.reduce((sum, discipline) => sum + discipline.encaminhamentos, 0),
      highlightedIssues
    };
  }, [visibleIssues, interfacesDisciplineFilter, interfacesStatusFilter, interfacesImpactFilter]);

  const interfacesPackageDashboard = useMemo(() => {
    const formatPackageDate = (value) => {
      const date = parseIssueDate(value);
      return date ? formatDate(date) : '-';
    };
    const sortByRecentEmission = (first, second) => {
      const firstDate = parseIssueDate(first.emittedDate || first.updatedAt || first.versionCreatedAt);
      const secondDate = parseIssueDate(second.emittedDate || second.updatedAt || second.versionCreatedAt);
      return (secondDate?.getTime() || 0) - (firstDate?.getTime() || 0);
    };

    const emittedRows = documentListRows.filter((document) => document.emitted);
    const pendingRows = documentListRows.filter((document) => !document.emitted);
    const packageMap = new Map();
    const disciplineMap = new Map();

    for (const document of documentListRows) {
      const discipline = document.discipline || 'Sem disciplina';
      const disciplineGroup = disciplineMap.get(discipline) || { discipline, total: 0, emitted: 0, pending: 0, packages: new Set(), latestDate: null };
      disciplineGroup.total += 1;
      if (document.emitted) disciplineGroup.emitted += 1;
      else disciplineGroup.pending += 1;
      if (document.emittedGrd || document.emittedFolder) disciplineGroup.packages.add(document.emittedGrd || document.emittedFolder);
      const documentDate = parseIssueDate(document.emittedDate || document.updatedAt || document.versionCreatedAt);
      if (documentDate && (!disciplineGroup.latestDate || documentDate > disciplineGroup.latestDate)) disciplineGroup.latestDate = documentDate;
      disciplineMap.set(discipline, disciplineGroup);

      if (!document.emitted) continue;
      const packageKey = document.emittedGrd || document.grd || document.emittedFolder || 'Sem GRD';
      const packageGroup = packageMap.get(packageKey) || {
        key: packageKey,
        documents: 0,
        disciplines: new Set(),
        revisions: new Set(),
        latestDate: null
      };
      packageGroup.documents += 1;
      if (discipline) packageGroup.disciplines.add(discipline);
      if (document.emittedRevision || document.revision) packageGroup.revisions.add(document.emittedRevision || document.revision);
      const emittedDate = parseIssueDate(document.emittedDate || document.updatedAt || document.versionCreatedAt);
      if (emittedDate && (!packageGroup.latestDate || emittedDate > packageGroup.latestDate)) packageGroup.latestDate = emittedDate;
      packageMap.set(packageKey, packageGroup);
    }

    const packages = Array.from(packageMap.values())
      .map((group) => ({
        ...group,
        disciplineCount: group.disciplines.size,
        revisionCount: group.revisions.size,
        latestDateLabel: group.latestDate ? formatDate(group.latestDate) : '-'
      }))
      .sort((first, second) => (second.latestDate?.getTime() || 0) - (first.latestDate?.getTime() || 0))
      .slice(0, 6);

    const revisions = emittedRows
      .filter((document) => document.emittedRevision || document.revision || document.emittedVersion)
      .sort(sortByRecentEmission)
      .slice(0, 6)
      .map((document) => ({
        id: document.id,
        code: document.code || document.emittedFileName || '-',
        title: document.title || document.emittedFileName || '-',
        discipline: document.discipline || 'Sem disciplina',
        revision: document.emittedRevision || document.revision || '-',
        version: document.emittedVersion || document.version || '-',
        grd: document.emittedGrd || document.grd || '-',
        date: formatPackageDate(document.emittedDate || document.updatedAt || document.versionCreatedAt),
        webView: document.webView
      }));

    const disciplines = Array.from(disciplineMap.values())
      .map((group) => ({
        ...group,
        packageCount: group.packages.size,
        latestDateLabel: group.latestDate ? formatDate(group.latestDate) : '-',
        percent: group.total ? Math.round((group.emitted / group.total) * 100) : 0
      }))
      .sort((first, second) => second.pending - first.pending || String(first.discipline).localeCompare(String(second.discipline), 'pt-BR', { sensitivity: 'base' }))
      .slice(0, 6);

    return {
      total: documentListRows.length,
      emitted: emittedRows.length,
      pending: pendingRows.length,
      packageCount: packageMap.size,
      revisionCount: revisions.length,
      packages,
      revisions,
      disciplines,
      hasData: documentListRows.length > 0
    };
  }, [documentListRows]);

  const selectedIssue = useMemo(
    () => issues.find((issue) => issue.id === selectedIssueId),
    [issues, selectedIssueId]
  );

  const selectedIssueFull = selectedIssueDetails || selectedIssue;

  const selectedIssueRows = useMemo(() => {
    if (!selectedIssueFull) return [];

    return [
      ['Status', selectedIssueFull.status],
      ['Prazo', selectedIssueFull.dueDate ? formatDate(selectedIssueFull.dueDate) : 'Sem prazo'],
      ['Categoria', selectedIssueFull.category || 'Sem categoria'],
      ['Tipo', selectedIssueFull.issueType || 'Sem tipo'],
      ['Subtipo', selectedIssueFull.issueSubtype || 'Sem subtipo'],
      ['Responsavel', selectedIssueFull.assignedTo],
      ['Criada em', formatDateTime(selectedIssueFull.createdAt)],
      ['Criada por', selectedIssueFull.createdBy],
      ['Atualizada em', formatDateTime(selectedIssueFull.updatedAt)],
      ['Atualizada por', selectedIssueFull.updatedBy],
      ['Aberta em', formatDateTime(selectedIssueFull.openedAt)],
      ['Fechada em', formatDateTime(selectedIssueFull.closedAt)],
      ['Localizacao', selectedIssueFull.location],
      ['Responsavel interno', selectedIssueFull.owner],
      ['Publicada', selectedIssueFull.published],
      ['Comentarios', selectedIssueFull.comments?.length ?? selectedIssueFull.commentCount],
      ['Anexos', selectedIssueFull.attachmentCount]
    ].filter(([, value]) => value !== undefined && value !== null && value !== '');
  }, [selectedIssueFull]);

  const issueLinkFlows = useMemo(() => {
    const issueLookup = buildIssueLookupIndexes(issues);
    const publishedIssues = issues.filter((issue) => issue.published !== false && normalizeText(issue.status) !== 'draft');
    const cache = new Map();
    const resolving = new Set();

    const resolveReferenceChain = (issue, localVisited = new Set()) => {
      if (!issue?.id || localVisited.has(issue.id)) return [];
      localVisited.add(issue.id);
      const directReferences = resolveIssueReferences(issue, issueLookup);
      const chain = [];
      const seen = new Set();

      for (const directReference of directReferences) {
        if (!directReference?.id || seen.has(directReference.id)) continue;
        seen.add(directReference.id);
        chain.push(directReference);

        for (const nestedReference of resolveReferenceChain(directReference, localVisited)) {
          if (!nestedReference?.id || seen.has(nestedReference.id)) continue;
          seen.add(nestedReference.id);
          chain.push(nestedReference);
        }
      }

      localVisited.delete(issue.id);
      return chain;
    };

    const resolveEffectiveDate = (issue) => {
      if (!issue) return null;
      if (cache.has(issue.id)) return cache.get(issue.id);
      if (resolving.has(issue.id)) return null;

      resolving.add(issue.id);
      let value = issue.dueDate || null;
      if (!value) {
        const refs = resolveReferenceChain(issue);
        for (const referenceIssue of refs) {
          const inherited = resolveEffectiveDate(referenceIssue);
          if (inherited) {
            value = inherited;
            break;
          }
        }
      }

      resolving.delete(issue.id);
      cache.set(issue.id, value);
      return value;
    };

    return publishedIssues
      .map((issue) => {
        const references = resolveReferenceChain(issue).sort((firstReference, secondReference) =>
          sortIssueHierarchy(firstReference, secondReference)
        );
        const effectiveDate = resolveEffectiveDate(issue);
        const parentIssue = references[0] || null;
        const discipline = getIssueDiscipline(issue) || (parentIssue ? getIssueDiscipline(parentIssue) : 'Sem disciplina');
        return { issue: { ...issue, discipline, effectiveDate }, references };
      })
      .sort((firstItem, secondItem) => {
        const a = firstItem.issue.effectiveDate ? new Date(firstItem.issue.effectiveDate).getTime() : Number.POSITIVE_INFINITY;
        const b = secondItem.issue.effectiveDate ? new Date(secondItem.issue.effectiveDate).getTime() : Number.POSITIVE_INFINITY;
        if (a !== b) return a - b;
        return sortByName(firstItem.issue, secondItem.issue);
      });
  }, [issues]);

  const linkCategoryOptions = useMemo(() => {
    return [...new Set(issues.map((issue) => issue.category || 'Sem categoria'))].sort((firstItem, secondItem) =>
      String(firstItem).localeCompare(String(secondItem), 'pt-BR', { sensitivity: 'base' })
    );
  }, [issues]);

  const linkFilterOptions = useMemo(() => {
    const build = (values) => [...new Set(values.filter(Boolean))].sort((a, b) => String(a).localeCompare(String(b), 'pt-BR', { sensitivity: 'base' }));
    return {
      categories: build(issues.map((issue) => issue.category || 'Sem categoria')),
      types: build(issues.map((issue) => issue.issueType || 'Sem tipo')),
      codigosMarco: build(issues.map((issue) => getCustomFieldValue(issue, customFieldDefinitions, 'Código do Marco') || 'Sem Código do Marco')),
      responsibles: build(issues.map((issue) => issue.assignedTo || 'Nao informado')),
      customFields: build(customFieldDefinitions.map((field) => field.title || field.name || field.id)),
      packages: build(issues.map((issue) => getCustomFieldValue(issue, customFieldDefinitions, 'Pacote/Marco Contratual') || 'Sem Marco Contratual'))
    };
  }, [issues, customFieldDefinitions]);

  const filteredLinkFlows = useMemo(() => {
    return issues.filter((issue) => {
      if (linkCategoryFilter !== 'all' && (issue.category || 'Sem categoria') !== linkCategoryFilter) return false;
      if (linkTypeFilter !== 'all' && (issue.issueType || 'Sem tipo') !== linkTypeFilter) return false;
      if (linkStatusFilter !== 'all' && normalizeText(issue.status) !== normalizeText(linkStatusFilter)) return false;
      const codigoMarco = getCustomFieldValue(issue, customFieldDefinitions, 'Código do Marco') || 'Sem Código do Marco';
      if (linkMarcoFilter !== 'all' && codigoMarco !== linkMarcoFilter) return false;
      if (linkResponsibleFilter !== 'all' && (issue.assignedTo || 'Nao informado') !== linkResponsibleFilter) return false;
      if (linkCustomFieldFilter !== 'all') {
        const fieldValue = getIssueCustomValue(issue, [linkCustomFieldFilter]);
        if (!normalizeText(fieldValue)) return false;
      }
      return true;
    });
  }, [issues, customFieldDefinitions, linkCategoryFilter, linkTypeFilter, linkStatusFilter, linkMarcoFilter, linkResponsibleFilter, linkCustomFieldFilter]);

  const isLinkFilterActive = useMemo(() => (
    linkCategoryFilter !== 'all' ||
    linkTypeFilter !== 'all' ||
    linkStatusFilter !== 'all' ||
    linkMarcoFilter !== 'all' ||
    linkResponsibleFilter !== 'all' ||
    linkCustomFieldFilter !== 'all'
  ), [linkCategoryFilter, linkTypeFilter, linkStatusFilter, linkMarcoFilter, linkResponsibleFilter, linkCustomFieldFilter]);

  const issueLinkGroups = useMemo(() => {
    const sourceIssues = filteredLinkFlows;
    const milestones = sourceIssues.filter((issue) => isDeliveryMilestoneIssue(issue, customFieldDefinitions));

    const sortFlowIssues = (flows) => flows.sort((firstItem, secondItem) => {
      const firstDate = getIssueTimelineDate(firstItem) || Number.POSITIVE_INFINITY;
      const secondDate = getIssueTimelineDate(secondItem) || Number.POSITIVE_INFINITY;
      if (firstDate !== secondDate) return firstDate - secondDate;
      return sortIssueHierarchy(firstItem, secondItem);
    });

    if (!milestones.length) {
      const fallbackGroups = new Map();

      for (const issue of sourceIssues) {
        const pack = getCustomFieldValue(issue, customFieldDefinitions, 'Código do Marco') || 'Sem Código do Marco';
        const group = fallbackGroups.get(pack) || { category: pack, milestone: null, dependencies: [], flows: [] };
        group.flows.push(issue);
        fallbackGroups.set(pack, group);
      }

      return Array.from(fallbackGroups.values())
        .map((group) => {
          const flows = sortFlowIssues(group.flows);
          return {
            ...group,
            dependencies: flows,
            total: flows.length,
            linked: flows.filter((issue) => (issue.relationshipReferenceIds || []).length > 0).length,
            open: flows.filter((issue) => isOpenIssue(issue)).length,
            closed: flows.filter((issue) => !isOpenIssue(issue)).length,
            overdue: flows.filter((issue) => isOverdue(issue)).length,
            progress: flows.length ? Math.round((flows.filter((issue) => !isOpenIssue(issue)).length / flows.length) * 100) : 0,
            flows
          };
        })
        .sort((firstItem, secondItem) => String(firstItem.category).localeCompare(String(secondItem.category), 'pt-BR', { sensitivity: 'base' }));
    }

    return milestones
      .map((milestone) => {
        const milestoneTokens = getIssueReferenceTokens(milestone);
        const dependencies = sortFlowIssues(sourceIssues
          .filter((issue) => issue.id !== milestone.id)
          .filter((issue) => {
            const issueTokens = getIssueReferenceTokens(issue);
            const issueReferencesMilestone = issueMatchesReferenceTokens(milestone, issueTokens);
            const milestoneReferencesIssue = issueMatchesReferenceTokens(issue, milestoneTokens);
            return issueReferencesMilestone || milestoneReferencesIssue;
          }));
        const flows = [milestone, ...dependencies].filter((issue, index, array) => array.findIndex((item) => item.id === issue.id) === index);
        const closedDependencies = dependencies.filter((issue) => !isOpenIssue(issue)).length;

        return {
          category: getMilestoneCode(milestone, customFieldDefinitions) || getMilestoneLabel(milestone, customFieldDefinitions),
          milestone,
          dependencies,
          flows,
          total: dependencies.length,
          linked: dependencies.length,
          open: dependencies.filter((issue) => isOpenIssue(issue)).length,
          closed: closedDependencies,
          overdue: dependencies.filter((issue) => isOverdue(issue)).length,
          progress: dependencies.length ? Math.round((closedDependencies / dependencies.length) * 100) : (!isOpenIssue(milestone) ? 100 : 0),
          isEmpty: dependencies.length === 0
        };
      })
      .sort((firstItem, secondItem) => String(firstItem.category).localeCompare(String(secondItem.category), 'pt-BR', { sensitivity: 'base' }));
  }, [filteredLinkFlows, customFieldDefinitions]);

  const kanbanIssueTypes = useMemo(
    () => ['Compatibilização', 'Decisão de Projeto', 'Definição Técnica Pendente', 'Solicitação de Informação'],
    []
  );

  const kanbanDisciplineOptions = useMemo(() => {
    const definition = findCustomFieldDefinition(customFieldDefinitions, ['Disciplina envolvida', 'Disciplina']);
    return (definition?.options || [])
      .map((option) => ({
        id: option.id || option.value || option.label,
        label: option.label || option.name || option.title || option.value || option.id
      }))
      .filter((option) => option.id && option.label)
      .sort((firstItem, secondItem) => String(firstItem.label).localeCompare(String(secondItem.label), 'pt-BR', { sensitivity: 'base' }));
  }, [customFieldDefinitions]);

  const linkKanbanColumns = useMemo(() => {
    const normalizeKanbanMatch = normalizeKanbanColumnKey;
    const isEntregaTitleIssue = (issue) => /^entrega\s*-/i.test(String(issue?.title || '').trim());
    const getKanbanMarcoSortParts = (code) => {
      const text = String(code || '').toUpperCase();
      const numbers = (text.match(/\d+/g) || []).map(Number);
      if (!numbers.length) return [9999, 9999, text];
      return [numbers[0], numbers[1] || 0, text];
    };
    const compareKanbanColumns = (firstItem, secondItem) => {
      const firstParts = getKanbanMarcoSortParts(firstItem.code);
      const secondParts = getKanbanMarcoSortParts(secondItem.code);
      if (firstParts[0] !== secondParts[0]) return firstParts[0] - secondParts[0];
      if (firstParts[1] !== secondParts[1]) return firstParts[1] - secondParts[1];
      return String(firstItem.marcoContratual || firstItem.title || firstItem.code).localeCompare(
        String(secondItem.marcoContratual || secondItem.title || secondItem.code),
        'pt-BR',
        { numeric: true, sensitivity: 'base' }
      );
    };
    const addIndexValue = (index, key, milestoneId) => {
      if (!key) return;
      if (!index.has(key)) {
        index.set(key, milestoneId);
        return;
      }
      if (index.get(key) !== milestoneId) {
        index.set(key, null);
      }
    };
    const readCardMarcoContratual = (issue) =>
      getCustomFieldValueByAliases(issue, customFieldDefinitions, MARCO_CONTRATUAL_FIELD_ALIASES) ||
      getIssueCustomValue(issue, MARCO_CONTRATUAL_FIELD_ALIASES) ||
      getCustomFieldValueByAliases(issue, customFieldDefinitions, KANBAN_MARCO_TITLE_FIELD_ALIASES) ||
      getIssueCustomValue(issue, KANBAN_MARCO_TITLE_FIELD_ALIASES) ||
      getCustomFieldValue(issue, customFieldDefinitions, 'Pacote/Marco Contratual') ||
      '';
    const readCardCodigoMarco = (issue) =>
      getCustomFieldValueByAliases(issue, customFieldDefinitions, CODIGO_MARCO_FIELD_ALIASES) ||
      getIssueCustomValue(issue, CODIGO_MARCO_FIELD_ALIASES) ||
      getCustomFieldValue(issue, customFieldDefinitions, 'Código do Marco') ||
      '';
    const getMilestoneMatchKeys = (milestone) => {
      const code = String(milestone.code || '').trim();
      const deliveryTitle = String(milestone.deliveryTitle || '').trim();
      const marcoContratual = String(milestone.marcoContratual || '').trim();
      const title = String(milestone.title || '').trim();
      return [
        deliveryTitle,
        marcoContratual,
        title,
        code,
        code && deliveryTitle ? `${code} ${deliveryTitle}` : '',
        code && deliveryTitle ? `${code} - ${deliveryTitle}` : '',
        code && marcoContratual ? `${code} ${marcoContratual}` : '',
        code && marcoContratual ? `${code} - ${marcoContratual}` : '',
        code && title ? `${code} ${title}` : '',
        code && title ? `${code} - ${title}` : ''
      ]
        .map(normalizeKanbanMatch)
        .filter(Boolean);
    };

    const milestones = issues
      .filter((issue) => isEntregaTitleIssue(issue) || isDeliveryMilestoneIssue(issue, customFieldDefinitions))
      .map((issue) => {
        const code = String(readCardCodigoMarco(issue) || getMilestoneCode(issue, customFieldDefinitions) || '').trim();
        const rawMarcoContratual = String(readCardMarcoContratual(issue) || '').trim();
        const deliveryTitle = String(issue.title || '').trim();
        const normalizedTitle = isEntregaTitleIssue(issue) ? deliveryTitle : '';
        const title = normalizedTitle || rawMarcoContratual || deliveryTitle || 'Entrega sem título';
        return {
          id: issue.id,
          issue,
          code,
          title,
          deliveryTitle: title,
          marcoContratual: title,
          marcoFieldValue: rawMarcoContratual || title,
          canReceiveCards: true
        };
      });

    const milestoneById = new Map(milestones.map((milestone) => [milestone.id, milestone]));
    const milestoneCodeIndex = new Map();
    const milestoneTitleIndex = new Map();
    milestones.forEach((milestone) => {
      addIndexValue(milestoneCodeIndex, normalizeKanbanMatch(milestone.code), milestone.id);
      getMilestoneMatchKeys(milestone).forEach((key) => addIndexValue(milestoneTitleIndex, key, milestone.id));
    });

    const isKanbanCardCandidate = (issue) => {
      if (isDeliveryMilestoneIssue(issue, customFieldDefinitions)) return false;
      return true;
    };

    const kanbanSourceIssues = isLinkFilterActive ? filteredLinkFlows : issues;
    const cards = kanbanSourceIssues.filter(isKanbanCardCandidate);
    const getCardMarcoData = (issue) => {
      const draftedColumnId = kanbanMoveDrafts[issue.id]?.columnId;
      if (draftedColumnId && milestoneById.has(draftedColumnId)) {
        const draftedColumn = milestoneById.get(draftedColumnId);
        return {
          code: draftedColumn.code || '',
          marcoContratual: draftedColumn.marcoContratual || draftedColumn.title || '',
          key: normalizeKanbanMatch(draftedColumn.marcoContratual || draftedColumn.title || draftedColumn.code)
        };
      }
      const code = String(readCardCodigoMarco(issue) || '').trim();
      const marcoContratual = String(readCardMarcoContratual(issue) || '').trim();
      return {
        code,
        marcoContratual,
        key: normalizeKanbanMatch(marcoContratual || code)
      };
    };

    const sortCards = (items) => [...items].sort((firstIssue, secondIssue) => {
      if (isOverdue(firstIssue) !== isOverdue(secondIssue)) return isOverdue(firstIssue) ? -1 : 1;
      const firstDate = firstIssue.dueDate ? new Date(firstIssue.dueDate).getTime() : Number.POSITIVE_INFINITY;
      const secondDate = secondIssue.dueDate ? new Date(secondIssue.dueDate).getTime() : Number.POSITIVE_INFINITY;
      if (firstDate !== secondDate) return firstDate - secondDate;
      return sortByName(firstIssue, secondIssue);
    });

    const columnsByKey = new Map();
    const cardsByColumnId = new Map();
    const registerColumn = (column) => {
      const key = normalizeKanbanMatch(column.deliveryTitle || column.title || column.marcoFieldValue || column.marcoContratual || column.code);
      if (!key) return null;
      const existing = columnsByKey.get(key);
      if (existing) {
        if (!existing.code && column.code) existing.code = column.code;
        if (!existing.issue && column.issue) existing.issue = column.issue;
        existing.canReceiveCards = existing.canReceiveCards || column.canReceiveCards;
        return existing;
      }
      const nextColumn = {
        ...column,
        id: column.id || `marco-${key}`,
        code: column.code || '',
        title: column.title || column.deliveryTitle || column.marcoFieldValue || column.marcoContratual || column.code || 'Entrega sem nome',
        deliveryTitle: column.deliveryTitle || column.title || column.marcoFieldValue || column.marcoContratual || column.code || 'Entrega sem nome',
        marcoContratual: column.marcoContratual || column.deliveryTitle || column.title || column.marcoFieldValue || column.code || 'Entrega sem nome',
        marcoFieldValue: column.marcoFieldValue || column.deliveryTitle || column.title || column.marcoContratual || column.code || 'Entrega sem nome',
        canReceiveCards: column.canReceiveCards !== false
      };
      columnsByKey.set(key, nextColumn);
      cardsByColumnId.set(nextColumn.id, []);
      milestoneById.set(nextColumn.id, nextColumn);
      if (nextColumn.code) addIndexValue(milestoneCodeIndex, normalizeKanbanMatch(nextColumn.code), nextColumn.id);
      getMilestoneMatchKeys(nextColumn).forEach((matchKey) => addIndexValue(milestoneTitleIndex, matchKey, nextColumn.id));
      return nextColumn;
    };

    milestones.forEach((milestone) => registerColumn(milestone));
    kanbanManualColumns.forEach((column) => registerColumn({
      ...column,
      issue: null,
      title: column.title || column.marcoContratual || column.code || 'Marco Contratual',
      deliveryTitle: column.deliveryTitle || column.title || column.marcoContratual || column.code || 'Marco Contratual',
      marcoContratual: column.marcoContratual || column.title || column.deliveryTitle || column.code || 'Marco Contratual',
      marcoFieldValue: column.marcoFieldValue || column.marcoContratual || column.title || column.code || 'Marco Contratual',
      canReceiveCards: true,
      isManual: true
    }));

    const unlinkedCards = [];
    const assignedCardIds = new Set();
    cards.forEach((issue) => {
      const cardId = issue.id || issue.displayId || issue.title;
      if (assignedCardIds.has(cardId)) return;
      assignedCardIds.add(cardId);

      const { code, marcoContratual, key } = getCardMarcoData(issue);
      let column = null;
      const titleKey = normalizeKanbanMatch(marcoContratual);
      const codeKey = normalizeKanbanMatch(code);

      if (titleKey && milestoneTitleIndex.has(titleKey)) {
        const milestoneId = milestoneTitleIndex.get(titleKey);
        column = milestoneId ? milestoneById.get(milestoneId) : null;
      }
      if (!column && codeKey && milestoneCodeIndex.has(codeKey)) {
        const milestoneId = milestoneCodeIndex.get(codeKey);
        column = milestoneId ? milestoneById.get(milestoneId) : null;
      }
      if (!column && key) {
        column = registerColumn({
          id: `marco-${key}`,
          issue: null,
          code: code || '',
          title: marcoContratual || code || 'Entrega sem nome',
          deliveryTitle: marcoContratual || code || 'Entrega sem nome',
          marcoContratual: marcoContratual || code || 'Entrega sem nome',
          marcoFieldValue: marcoContratual || code || 'Entrega sem nome',
          isDynamic: true,
          canReceiveCards: true
        });
      }

      if (column?.id && cardsByColumnId.has(column.id)) {
        cardsByColumnId.get(column.id).push(issue);
        return;
      }
      unlinkedCards.push(issue);
    });

    const columns = Array.from(columnsByKey.values())
      .sort(compareKanbanColumns)
      .map((column) => {
        const columnCards = sortCards(cardsByColumnId.get(column.id) || []);
        return {
          ...column,
          cards: columnCards,
          open: columnCards.filter(isOpenIssue).length,
          overdue: columnCards.filter(isOverdue).length,
          closed: columnCards.filter((issue) => !isOpenIssue(issue)).length
        };
      });

    const sortedUnlinkedCards = sortCards(unlinkedCards);
    if (sortedUnlinkedCards.length) {
      columns.push({
        id: 'sem-marco',
        issue: null,
        code: 'SEM MARCO',
        title: 'Sem entrega vinculada',
        deliveryTitle: 'Sem entrega vinculada',
        marcoContratual: 'Sem entrega vinculada',
        marcoFieldValue: 'Sem entrega vinculada',
        canReceiveCards: false,
        cards: sortedUnlinkedCards,
        open: sortedUnlinkedCards.filter(isOpenIssue).length,
        overdue: sortedUnlinkedCards.filter(isOverdue).length,
        closed: sortedUnlinkedCards.filter((issue) => !isOpenIssue(issue)).length
      });
    }

    return columns;
  }, [issues, filteredLinkFlows, customFieldDefinitions, isLinkFilterActive, kanbanMoveDrafts, kanbanManualColumns]);

  const selectedKanbanIssue = useMemo(
    () => issues.find((issue) => issue.id === selectedKanbanIssueId) || null,
    [issues, selectedKanbanIssueId]
  );

  const selectedKanbanColumn = useMemo(
    () => linkKanbanColumns.find((column) => column.cards.some((issue) => issue.id === selectedKanbanIssueId)) || null,
    [linkKanbanColumns, selectedKanbanIssueId]
  );

  useEffect(() => {
    if (!selectedKanbanIssueId) return;
    if (!selectedKanbanIssue) {
      setSelectedKanbanIssueId('');
    }
  }, [selectedKanbanIssueId, selectedKanbanIssue]);

  useEffect(() => {
    if (!selectedKanbanIssue) {
      setKanbanEditDraft({
        marcoId: '',
        title: '',
        assignedTo: '',
        followers: [],
        dueDate: '',
        disciplina: '',
        description: '',
        comment: ''
      });
      return;
    }

    setKanbanEditDraft({
      marcoId: selectedKanbanColumn?.canReceiveCards ? selectedKanbanColumn.id : '',
      title: selectedKanbanIssue.title || '',
      assignedTo: getKanbanAssignedUserId(selectedKanbanIssue),
      followers: getKanbanFollowerIds(selectedKanbanIssue),
      dueDate: getDateInputValue(selectedKanbanIssue.dueDate),
      disciplina: getKanbanDisciplineOptionValue(selectedKanbanIssue),
      description: selectedKanbanIssue.description || '',
      comment: ''
    });
  }, [selectedKanbanIssue, selectedKanbanColumn?.id, kanbanDisciplineOptions, projectUsers]);

  const selectedImpactFieldLabel = impactFields.find((field) => field.value === selectedImpactField)?.label || impactFields[0].label;
  const usesImpactOptionValues = ['nivel-impacto', 'impacto-escopo', 'fase'].includes(selectedImpactField);

  const impactTypeOptions = useMemo(
    () => [...new Set(issues.map((issue) => issue.issueType || 'Sem tipo'))].sort((firstType, secondType) =>
      String(firstType).localeCompare(String(secondType), 'pt-BR', { sensitivity: 'base' })
    ),
    [issues]
  );

  const impactCategoryOptions = useMemo(
    () => [...new Set(issues.map((issue) => issue.category || 'Sem categoria'))].sort((firstCategory, secondCategory) =>
      String(firstCategory).localeCompare(String(secondCategory), 'pt-BR', { sensitivity: 'base' })
    ),
    [issues]
  );

  const impactPhaseOptions = useMemo(
    () =>
      [...new Set(issues.map((issue) => getImpactFieldValue(issue, 'fase')).filter((value) => value && value !== 'Nao informado'))].sort(
        (firstPhase, secondPhase) => String(firstPhase).localeCompare(String(secondPhase), 'pt-BR', { sensitivity: 'base' })
      ),
    [issues]
  );

  const impactFilteredSourceIssues = useMemo(() => {
    const startTime = impactStartDate ? new Date(`${impactStartDate}T00:00:00.000Z`).getTime() : null;
    const endTime = impactEndDate ? new Date(`${impactEndDate}T23:59:59.999Z`).getTime() : null;

    return issues.filter((issue) => {
      const issueType = issue.issueType || 'Sem tipo';
      const issueCategory = issue.category || 'Sem categoria';

      if (impactCategoryFilter !== 'all' && issueCategory !== impactCategoryFilter) return false;
      if (impactTypeFilter !== 'all' && issueType !== impactTypeFilter) return false;

      const issuePhase = getImpactFieldValue(issue, 'fase');
      if (impactPhaseFilter !== 'all' && normalizeText(issuePhase) !== normalizeText(impactPhaseFilter)) return false;

      if (startTime || endTime) {
        if (!issue.dueDate) return false;
        const dueTime = new Date(issue.dueDate).getTime();
        if (!Number.isFinite(dueTime)) return false;
        if (startTime && dueTime < startTime) return false;
        if (endTime && dueTime > endTime) return false;
      }

      return true;
    });
  }, [issues, impactStartDate, impactEndDate, impactTypeFilter, impactCategoryFilter, impactPhaseFilter]);

  const impactIssues = useMemo(() => {
    return impactFilteredSourceIssues
      .map((issue) => {
        const rawImpact = getImpactFieldValue(issue, selectedImpactField);
        const impactLevel = usesImpactOptionValues ? normalizeImpactOption(rawImpact) : normalizeImpactLevel(rawImpact);
        return {
          ...issue,
          impactRawValue: rawImpact,
          impactLevel,
          impactLevelKey: normalizeText(impactLevel).replace(/\s+/g, '-'),
          dueGroup: issue.dueDate ? getMonthLabel(getMonthKey(issue.dueDate)) : 'Sem prazo definido'
        };
      })
      .sort((firstIssue, secondIssue) => {
        if (!firstIssue.dueDate && secondIssue.dueDate) return 1;
        if (firstIssue.dueDate && !secondIssue.dueDate) return -1;
        if (firstIssue.dueDate && secondIssue.dueDate) {
          const dateResult = new Date(firstIssue.dueDate) - new Date(secondIssue.dueDate);
          if (dateResult !== 0) return dateResult;
        }
        return sortByName(firstIssue, secondIssue);
      });
  }, [impactFilteredSourceIssues, selectedImpactField, usesImpactOptionValues]);

  const activeImpactLevels = useMemo(() => {
    if (!usesImpactOptionValues) return impactLevels;

    const levelsByKey = new Map();

    for (const issue of impactIssues) {
      const label = issue.impactLevel || 'Sem classificação';
      const key = normalizeText(label).replace(/[^a-z0-9]+/g, '-').replace(/(^-|-$)/g, '') || 'sem-classificacao';

      if (!levelsByKey.has(key)) {
        levelsByKey.set(key, {
          value: label,
          key,
          label,
          color: label === 'Sem classificação' ? '#7b8290' : impactOptionPalette[levelsByKey.size % impactOptionPalette.length]
        });
      }
    }

    if (!levelsByKey.has('sem-classificacao')) {
      levelsByKey.set('sem-classificacao', {
        value: 'Sem classificação',
        key: 'sem-classificacao',
        label: 'Sem classificação',
        color: '#7b8290'
      });
    }

    return [...levelsByKey.values()].sort((firstLevel, secondLevel) => {
      if (firstLevel.value === 'Sem classificação') return 1;
      if (secondLevel.value === 'Sem classificação') return -1;
      return String(firstLevel.label).localeCompare(String(secondLevel.label), 'pt-BR', { sensitivity: 'base' });
    });
  }, [impactIssues, usesImpactOptionValues]);

  const impactMetrics = useMemo(() => {
    const countByLevel = new Map(activeImpactLevels.map((level) => [level.value, 0]));

    for (const issue of impactIssues) {
      countByLevel.set(issue.impactLevel, (countByLevel.get(issue.impactLevel) || 0) + 1);
    }

    return {
      total: impactIssues.length,
      withoutDueDate: impactIssues.filter((issue) => !issue.dueDate).length,
      byLevel: activeImpactLevels.map((level) => ({
        ...level,
        total: countByLevel.get(level.value) || 0
      }))
    };
  }, [impactIssues, activeImpactLevels]);

  const impactTimelineGroups = useMemo(() => {
    const groups = new Map();

    for (const issue of impactIssues) {
      const key = issue.dueDate ? getMonthKey(issue.dueDate) : 'sem-prazo';
      const group = groups.get(key) || {
        key,
        label: issue.dueDate ? getMonthLabel(key) : 'Sem prazo definido',
        issues: []
      };
      group.issues.push(issue);
      groups.set(key, group);
    }

    return [...groups.values()].sort((firstGroup, secondGroup) => {
      if (firstGroup.key === 'sem-prazo') return 1;
      if (secondGroup.key === 'sem-prazo') return -1;
      return String(firstGroup.key).localeCompare(String(secondGroup.key));
    });
  }, [impactIssues]);

  const impactTimelineData = useMemo(() => {
    const datedIssues = impactIssues.filter((issue) => issue.dueDate);
    const dateValues = datedIssues
      .map((issue) => new Date(issue.dueDate).getTime())
      .filter((value) => Number.isFinite(value))
      .sort((firstDate, secondDate) => firstDate - secondDate);
    const uniqueTicks = [...new Set(dateValues.map((value) => new Date(value).toISOString().slice(0, 10)))];
    const positionByDate = new Map(
      uniqueTicks.map((dateKey, index) => [
        dateKey,
        uniqueTicks.length > 1 ? 4 + (index / (uniqueTicks.length - 1)) * 92 : 50
      ])
    );
    const visibleTickKeys =
      uniqueTicks.length <= 8
        ? uniqueTicks
        : uniqueTicks.filter((_, index) => {
            const interval = Math.ceil((uniqueTicks.length - 1) / 7);
            return index === 0 || index === uniqueTicks.length - 1 || index % interval === 0;
          });
    const ticks = visibleTickKeys.map((dateKey) => ({
      key: dateKey,
      label: formatDate(dateKey),
      position: positionByDate.get(dateKey) || 50
    }));
    const lanes = activeImpactLevels.map((level) => ({
      ...level,
      issues: impactIssues
        .filter((issue) => issue.impactLevel === level.value && issue.dueDate)
        .map((issue) => ({
          ...issue,
          position: positionByDate.get(new Date(issue.dueDate).toISOString().slice(0, 10)) || 50
        }))
    }));

    return {
      ticks,
      lanes,
      withoutDueDate: impactIssues.filter((issue) => !issue.dueDate)
    };
  }, [impactIssues, activeImpactLevels]);

  const projectManagementRows = useMemo(() => {
    const rowMap = new Map();
    const readCustom = (issue, aliases) => emptyIfNotInformed(getCustomFieldValueByAliases(issue, customFieldDefinitions, aliases));
    const calculateExecutionPercent = (summary) => {
      if (!summary?.total) return 0;
      return Math.round((summary.concluidas / summary.total) * 100);
    };
    const getStatusBucket = (issue) => {
      const normalizedStatus = normalizeText(issue.status || issue.raw?.attributes?.status);
      if (['closed', 'completed', 'done', 'concluido', 'fechado', 'finalizado', 'aprovado'].some((status) => normalizedStatus.includes(status))) return 'concluidas';
      if (['in_progress', 'em andamento', 'andamento', 'progress'].some((status) => normalizedStatus.includes(status))) return 'andamento';
      if (['pending', 'pendente', 'aguardando'].some((status) => normalizedStatus.includes(status))) return 'pendentes';
      if (['open', 'aberto'].some((status) => normalizedStatus.includes(status))) return 'abertas';
      return 'outros';
    };

    for (const issue of issues) {
      const milestoneCode =
        getMilestoneCode(issue, customFieldDefinitions) ||
        readCustom(issue, ['Marco de Medição', 'Marco Medicao', 'Marco de Medicao', 'Marco', 'Código do Marco', 'Codigo do Marco']);
      const measurementValue = readCustom(issue, [
        'Valor da Medição',
        'Valor da Medicao',
        'Valor Medicao',
        'Valor Medição'
      ]);
      const value = measurementValue || readCustom(issue, [
        'Valor',
        'Valor do Marco',
        'Valor Contratual',
        'Valor Medido'
      ]);
      const measurementStatus = readCustom(issue, [
        'Status medição',
        'Status da Medição',
        'Status Medicao',
        'Status da Medicao',
        'Status Vale'
      ]);
      const hasManagementSignal = milestoneCode || value || measurementStatus || normalizeText(issue.issueType).includes('marco contratual');

      if (!hasManagementSignal) continue;

      const key = normalizeText(milestoneCode || issue.displayId || issue.title || issue.id);
      const deliverable =
        readCustom(issue, ['Entregável', 'Entregavel', 'Atividade', 'Descrição do Entregável', 'Descricao do Entregavel']) ||
        issue.title ||
        'Sem entregável informado';
      const contractualDate =
        readCustom(issue, ['Prazo contratual', 'Data contratual']) ||
        '';
      const expectedDate =
        readCustom(issue, ['Data prevista G5', 'Prazo de resposta']) ||
        issue.dueDate ||
        '';
      const progress = formatPercentValue(readCustom(issue, ['Avanço', 'Avanço físico', 'Avanco fisico', '% Avanço Físico', '% Avanco Fisico', 'Execução', 'Execucao']));
      const risk = normalizeImpactLevel(
        readCustom(issue, [
          'Risco',
          'Nível de impacto',
          'Nível de impacto',
          'Nivel de impacto',
          'Impacto no escopo',
          'Impacto em escopo',
          'Impacto previsto',
          'Risco Previsto',
          'Tipo de Risco / Restrição',
          'Tipo de Risco / Restrição',
          'Tipo de Risco / Restricao',
          'Impacto no Cronograma',
          'Impacto em Prazo',
          'Impacto em Medição',
          'Impacto na Medicao'
        ]) ||
          (isOverdue(issue) ? 'Alto' : '')
      );
      const statusMedicao =
        measurementStatus ||
        readCustom(issue, ['Medição', 'Medicao']) ||
        (normalizeText(issue.status).includes('closed') || normalizeText(issue.status).includes('concluido') ? 'Medido' : 'A medir');
      const isMilestoneIssue = normalizeText(issue.issueType).includes('marco contratual');

      const existing = rowMap.get(key) || {
        key,
        marco: milestoneCode || issue.displayId || 'Sem marco',
        milestoneTitle: '',
        deliverables: [],
        issueItems: [],
        value: '',
        measurementValue: '',
        measurementIssueId: '',
        contractualDate: '',
        expectedDate: '',
        progress: '',
        risk: 'Sem classificação',
        statusMedicao: '',
        statusSummary: {
          total: 0,
          concluidas: 0,
          abertas: 0,
          andamento: 0,
          pendentes: 0,
          outros: 0
        },
        issues: []
      };

      if (!existing.milestoneTitle && isMilestoneIssue) existing.milestoneTitle = issue.title || deliverable;
      if (deliverable && !existing.deliverables.includes(deliverable)) existing.deliverables.push(deliverable);
      if (!existing.issueItems.some((item) => item.id === issue.id)) {
        existing.issueItems.push({
          id: issue.id,
          title: deliverable,
          status: issue.status || '',
          displayId: issue.displayId || ''
        });
      }
      if (!existing.value && value) existing.value = value;
      if (!existing.measurementValue && measurementValue) existing.measurementValue = measurementValue;
      if (!existing.measurementIssueId && (measurementValue || normalizeText(issue.issueType).includes('marco contratual'))) {
        existing.measurementIssueId = issue.id;
      }
      if (!existing.contractualDate && contractualDate) existing.contractualDate = contractualDate;
      if (!existing.expectedDate && expectedDate) existing.expectedDate = expectedDate;
      if (!existing.progress && progress) existing.progress = progress;
      if (sortRiskLevel(risk, existing.risk) < 0) existing.risk = risk;
      if (!existing.statusMedicao || normalizeText(existing.statusMedicao).includes('medido')) existing.statusMedicao = statusMedicao;
      existing.issues.push(issue);
      const statusBucket = getStatusBucket(issue);
      existing.statusSummary.total += 1;
      existing.statusSummary[statusBucket] += 1;
      if (!existing.measurementIssueId) existing.measurementIssueId = existing.issues[0]?.id || '';
      rowMap.set(key, existing);
    }

    const search = normalizeText(projectManagementSearch);
    return [...rowMap.values()]
      .map((row) => ({
        ...row,
        milestoneTitle: row.milestoneTitle || row.deliverables[0] || row.marco,
        entregavel: row.deliverables.join(' | ') || 'Sem entregável informado',
        valueInput: formatCurrencyBRL(row.measurementValue || row.value || 0),
        issueCount: row.issues.length,
        executionPercent: calculateExecutionPercent(row.statusSummary),
        progress: `${calculateExecutionPercent(row.statusSummary)}%`
      }))
      .filter((row) => {
        if (projectManagementStatusFilter !== 'all' && normalizeText(row.statusMedicao) !== normalizeText(projectManagementStatusFilter)) return false;
        if (!search) return true;
        return normalizeText([row.marco, row.entregavel, row.value, row.risk, row.statusMedicao].join(' ')).includes(search);
      })
      .sort((firstRow, secondRow) => String(firstRow.marco).localeCompare(String(secondRow.marco), 'pt-BR', { numeric: true, sensitivity: 'base' }));
  }, [issues, customFieldDefinitions, projectManagementSearch, projectManagementStatusFilter]);

  const projectManagementMetrics = useMemo(() => {
    const medidos = projectManagementRows.filter((row) => normalizeText(row.statusMedicao).includes('medido')).length;
    const aMedir = projectManagementRows.filter((row) => normalizeText(row.statusMedicao).includes('medir')).length;
    const highRisk = projectManagementRows.filter((row) => normalizeText(row.risk).includes('alto')).length;
    const visibleIssueIds = new Set(projectManagementRows.flatMap((row) => row.issues.map((issue) => issue.id)));
    const totalValue = projectManagementRows.reduce((sum, row) => sum + parseCurrencyNumber(row.measurementValue || row.value), 0);
    const measuredValue = projectManagementRows
      .filter((row) => normalizeText(row.statusMedicao).includes('medido'))
      .reduce((sum, row) => sum + parseCurrencyNumber(row.measurementValue || row.value), 0);
    const execution = projectManagementRows.length
      ? Math.round(projectManagementRows.reduce((sum, row) => sum + (row.executionPercent || 0), 0) / projectManagementRows.length)
      : 0;
    const statusData = [
      { name: 'Medidos', value: medidos, color: '#0f7c90' },
      { name: 'A medir', value: aMedir, color: '#f0a120' }
    ];
    const riskData = ['Alto', 'Médio', 'Baixo', 'Sem classificação'].map((risk) => ({
      name: risk,
      value: projectManagementRows.filter((row) => normalizeText(row.risk) === normalizeText(risk)).length,
      color: risk === 'Alto' ? '#d97373' : risk === 'Médio' ? '#f0a120' : risk === 'Baixo' ? '#4eb567' : '#8a97a0'
    }));
    const topMilestones = [...projectManagementRows]
      .sort((firstRow, secondRow) => secondRow.issueCount - firstRow.issueCount)
      .slice(0, 6);

    return {
      milestones: projectManagementRows.length,
      visibleIssues: visibleIssueIds.size,
      medidos,
      aMedir,
      highRisk,
      execution,
      totalValue,
      measuredValue,
      pendingValue: Math.max(0, totalValue - measuredValue),
      statusData,
      riskData,
      topMilestones
    };
  }, [projectManagementRows]);

  const projectManagementStatusOptions = useMemo(
    () => [...new Set(projectManagementRows.map((row) => row.statusMedicao).filter(Boolean))].sort((firstStatus, secondStatus) =>
      String(firstStatus).localeCompare(String(secondStatus), 'pt-BR', { sensitivity: 'base' })
    ),
    [projectManagementRows]
  );

  useEffect(() => {
    if (activeModule !== 'impact' || !selectedIssueId) return;
    if (!impactIssues.some((issue) => issue.id === selectedIssueId)) {
      setSelectedIssueId('');
    }
  }, [activeModule, impactIssues, selectedIssueId]);

  const documentMetrics = useMemo(
    () => ({
      total: documentListRows.length,
      emitted: documentListRows.filter((document) => document.emitted).length,
      pending: documentListRows.filter((document) => !document.emitted).length,
      withGrd: documentListRows.filter((document) => document.emittedGrd).length
    }),
    [documentListRows]
  );
  const documentEmittedPercentage = useMemo(() => {
    if (!documentMetrics.total) return 0;
    return Math.round((documentMetrics.emitted / documentMetrics.total) * 100);
  }, [documentMetrics.emitted, documentMetrics.total]);

  const bimAwpDashboard = useMemo(() => {
    const uniqueValues = (values) => [...new Set(values.map((value) => String(value || '').trim()).filter(Boolean))];
    const hasTerm = (issue, terms) => {
      const customText = Object.values(issue.customFields || issue.customAttributes || issue.custom_attributes || {})
        .map(toDisplayCustomFieldValue)
        .filter(Boolean)
        .join(' ');
      const source = normalizeText([
        issue.title,
        issue.description,
        issue.category,
        issue.issueType,
        issue.issueSubtype,
        customText
      ].filter(Boolean).join(' '));
      return terms.some((term) => source.includes(normalizeText(term)));
    };

    const related = visibleIssues.filter((issue) => hasTerm(issue, [
      'bim',
      'awp',
      'workface',
      'workfaces',
      'modelo',
      'interferencia',
      'interferência',
      'validacao',
      'validação',
      'qualidade',
      'publicacao',
      'publicação'
    ]));
    const workfaces = uniqueValues(related.map((issue) => getCustomFieldValueByAliases(issue, customFieldDefinitions, [
      'Workface',
      'Workface / Estrutura',
      'Pacote de Trabalho',
      'Pacote AWP',
      'CWP',
      'IWP'
    ]))).filter((value) => value && normalizeText(value) !== 'nao informado');
    const modelIssues = related.filter((issue) => hasTerm(issue, ['modelo', 'model', 'bim', 'workface', 'workfaces', 'awp']));
    const validationIssues = related.filter((issue) => hasTerm(issue, ['validacao', 'validação', 'qualidade', 'interferencia', 'interferência']));
    const publicationIssues = related.filter((issue) => hasTerm(issue, ['publicacao', 'publicação', 'publicado', 'emissao', 'emissão']));
    const statusGroups = new Map();
    related.forEach((issue) => {
      const status = issue.status || 'Sem status';
      statusGroups.set(status, (statusGroups.get(status) || 0) + 1);
    });

    return {
      related,
      total: related.length,
      open: related.filter(isOpenIssue).length,
      overdue: related.filter(isOverdue).length,
      modelIssues: modelIssues.length,
      validationIssues: validationIssues.length,
      publicationIssues: publicationIssues.length,
      workfaces: workfaces.length,
      statusGroups: [...statusGroups.entries()].map(([name, value]) => ({ name, value })),
      documentsTotal: documentMetrics.total,
      documentsEmitted: documentMetrics.emitted,
      documentsPending: documentMetrics.pending,
      documentsPercent: documentEmittedPercentage
    };
  }, [visibleIssues, customFieldDefinitions, documentMetrics, documentEmittedPercentage]);

  const documentSummary = useMemo(() => {
    const unique = (values) => new Set(values.map((value) => String(value || '').trim()).filter(Boolean)).size;
    return {
      disciplines: unique(documentListRows.map((document) => document.discipline)),
      documentTypes: unique(documentListRows.map((document) => document.documentType)),
      folders: unique(documentListRows.map((document) => document.emittedFolder))
    };
  }, [documentListRows]);

  const documentTypeOptions = useMemo(
    () =>
      [...new Set(documentListRows.map((document) => String(document.documentType || '').trim()).filter(Boolean))].sort((first, second) =>
        String(first).localeCompare(String(second), 'pt-BR', { sensitivity: 'base' })
      ),
    [documentListRows]
  );

  const filteredDocumentRows = useMemo(() => {
    const normalizedSearch = documentSearch
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .trim();

    return documentListRows.filter((document) => {
      if (documentStatusFilter === 'emitted' && !document.emitted) return false;
      if (documentStatusFilter === 'pending' && document.emitted) return false;
      if (documentTypeFilter !== 'all' && document.documentType !== documentTypeFilter) return false;

      if (!normalizedSearch) return true;

      const searchable = [
        document.code,
        document.title,
        document.documentType,
        document.discipline,
        document.emittedFileName,
        document.emittedGrd,
        document.emittedRevision,
        document.status,
        document.emittedDate,
        document.emissionDate
      ]
        .join(' ')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase();

      return searchable.includes(normalizedSearch);
    });
  }, [documentListRows, documentSearch, documentStatusFilter, documentTypeFilter]);

  const sortedDocumentRows = useMemo(() => {
    const direction = documentSort.direction === 'desc' ? -1 : 1;

    return [...filteredDocumentRows].sort((firstDocument, secondDocument) => {
      const firstValue = String(firstDocument[documentSort.field] || '');
      const secondValue = String(secondDocument[documentSort.field] || '');
      const result = String(firstValue).localeCompare(String(secondValue), 'pt-BR', { numeric: true, sensitivity: 'base' });
      if (result !== 0) return result * direction;
      return String(firstDocument.code || '').localeCompare(String(secondDocument.code || ''), 'pt-BR', {
        numeric: true,
        sensitivity: 'base'
      });
    });
  }, [filteredDocumentRows, documentSort]);

  const eapMetrics = useMemo(() => {
    const rows = eapEditedRows.length ? eapEditedRows : eapPreview?.rows || [];
    return {
      total: rows.length,
      ready: rows.filter((row) => row.validation === 'pronto').length,
      incomplete: rows.filter((row) => row.validation === 'incompleto' || row.validation === 'alerta').length,
      duplicated: rows.filter((row) => row.validation === 'duplicado').length,
      errors: rows.filter((row) => row.validation === 'erro').length,
      creatable: rows.filter((row) => row.validation !== 'erro' && row.validation !== 'duplicado').length,
      edited: Object.keys(eapEditedCells).length
    };
  }, [eapPreview, eapEditedRows, eapEditedCells]);

  const eapVisibleRows = useMemo(() => (eapEditedRows.length ? eapEditedRows : eapPreview?.rows || []).slice(0, eapPreviewLimit), [eapPreview, eapEditedRows]);

  const documentsByFolder = useMemo(() => {
    const groups = new Map();

    for (const document of publishedDocuments) {
      const folderName = document.folderPath || 'Sem pasta';
      const folderDocuments = groups.get(folderName) || [];
      folderDocuments.push(document);
      groups.set(folderName, folderDocuments);
    }

    return [...groups.entries()]
      .sort(([firstFolder], [secondFolder]) => String(firstFolder).localeCompare(String(secondFolder), 'pt-BR', { sensitivity: 'base' }))
      .map(([folderName, documents]) => ({
        folderName,
        documents: documents.sort((firstDocument, secondDocument) => sortByName(firstDocument, secondDocument))
      }));
  }, [publishedDocuments]);

  useEffect(() => {
    if (loginMessage) setError(loginMessage);
    loadInitialData();
  }, [loginMessage]);

  useEffect(() => {
    if (
      activeModule === 'interfaces-pendencias'
      && selectedProjectId
      && selectedHubId
      && !documentsLoading
      && documentListRows.length === 0
      && publishedDocuments.length === 0
    ) {
      loadPublishedDocuments(selectedProjectId);
    }
  }, [activeModule, selectedProjectId, selectedHubId]);

  useEffect(() => {
    const handleUnexpectedError = (event) => {
      setError(event.reason?.message || event.error?.message || event.message || 'Ocorreu um erro inesperado na tela.');
    };

    window.addEventListener('error', handleUnexpectedError);
    window.addEventListener('unhandledrejection', handleUnexpectedError);

    return () => {
      window.removeEventListener('error', handleUnexpectedError);
      window.removeEventListener('unhandledrejection', handleUnexpectedError);
    };
  }, []);


  useEffect(() => {
    if (!user) return undefined;

    const guardState = { source: 'app-guard' };
    window.history.pushState(guardState, '', window.location.href);

    const handlePopState = () => {
      window.history.pushState(guardState, '', window.location.href);
    };

    window.addEventListener('popstate', handlePopState);

    return () => {
      window.removeEventListener('popstate', handlePopState);
    };
  }, [user]);

  useEffect(() => {
    if (!user) return undefined;

    const keepSessionAlive = () => {
      fetch('/api/auth/keep-alive', {
        method: 'POST',
        credentials: 'include'
      }).catch(() => {});
    };

    keepSessionAlive();
    const intervalId = window.setInterval(keepSessionAlive, 5 * 60 * 1000);

    return () => window.clearInterval(intervalId);
  }, [user]);

  useEffect(() => {
    if (!selectedProjectId || !selectedIssueId) {
      setSelectedIssueDetails(null);
      setSelectedIssueDetailsError('');
      setSelectedIssueDetailsLoading(false);
      return undefined;
    }

    let cancelled = false;
    setSelectedIssueDetails(null);
    setSelectedIssueDetailsError('');
    setSelectedIssueDetailsLoading(true);

    requestJson(
      `/api/projects/${encodeURIComponent(selectedProjectId)}/issues/${encodeURIComponent(selectedIssueId)}/details`
    )
      .then((details) => {
        if (!cancelled) {
          setSelectedIssueDetails(details);
        }
      })
      .catch((requestError) => {
        if (!cancelled) {
          setSelectedIssueDetailsError(requestError.message);
        }
      })
      .finally(() => {
        if (!cancelled) {
          setSelectedIssueDetailsLoading(false);
        }
      });

    return () => {
      cancelled = true;
    };
  }, [selectedProjectId, selectedIssueId]);

  async function requestJson(url, options = {}) {
    const response = await fetch(url, {
      ...options,
      credentials: 'include',
      headers: {
        'Content-Type': 'application/json',
        ...(options.headers || {})
      }
    });

    if (!response.ok) {
      const text = await response.text();
      const contentType = response.headers.get('content-type') || '';
      const receivedHtml =
        contentType.includes('text/html') || text.trim().toLowerCase().startsWith('<!doctype html') || text.trim().startsWith('<');
      let data = {};

      try {
        data = text ? JSON.parse(text) : {};
      } catch {
        data = {};
      }

      const requestError = new Error(
        data.message ||
          (receivedHtml
            ? `O servidor retornou uma pagina de erro (${response.status}). Tente atualizar novamente em alguns segundos.`
            : text) ||
          'Nao foi possivel concluir a solicitacao.'
      );
      requestError.status = response.status;
      throw requestError;
    }

    return response.json();
  }

  async function loadInitialData() {
    setLoading(true);
    setError('');

    try {
      const me = await requestJson('/api/me');
      setUser(me);

      const hubsData = await requestJson('/api/hubs');
      setHubs(hubsData);
    } catch (requestError) {
      setUser(null);
      setHubs([]);

      if (!String(requestError.message).includes('Você precisa entrar')) {
        setError(requestError.message);
      }
    } finally {
      setLoading(false);
    }
  }

  function handleBusinessUnitSelect(businessUnit) {
    setSelectedBusinessUnit(businessUnit);
    localStorage.setItem('central-g5-business-unit', businessUnit);
    setSelectedHubId('');
    setSelectedProjectId('');
    setSelectedIssueId('');
    setProjects([]);
    setIssues([]);
    setIssueRelationshipsById({});
    setActiveModule('');
    setProjectPickerOpen(true);
    setError('');
  }

  function handleBusinessUnitReset() {
    setSelectedBusinessUnit('');
    localStorage.removeItem('central-g5-business-unit');
    setSelectedHubId('');
    setSelectedProjectId('');
    setSelectedIssueId('');
    setProjects([]);
    setIssues([]);
    setIssueRelationshipsById({});
    setActiveModule('');
    setProjectPickerOpen(false);
    setError('');
  }

  function updateInstrumentResource(resourceKey, field, value) {
    const cacheKey = `central-g5-instrumentos-recursos:${selectedProjectId}`;
    setInstrumentResourceOverrides((current) => {
      const next = {
        ...current,
        [resourceKey]: {
          ...(current[resourceKey] || {}),
          [field]: value
        }
      };
      localStorage.setItem(cacheKey, JSON.stringify(next));
      return next;
    });
  }

  async function handleHubChange(event) {
    const hubId = event.target.value;
    setSelectedHubId(hubId);
    setSelectedProjectId('');
    setSelectedIssueId('');
    setActiveModule('');
    setProjectPickerOpen(true);
    setIssueFilter('all');
    setCategoryFilter('all');
    setTypeFilter('all');
    setLinkCategoryFilter('all');
    setLinkTypeFilter('all');
    setLinkStatusFilter('all');
    setLinkMarcoFilter('all');
    setLinkResponsibleFilter('all');
    setLinkCustomFieldFilter('all');
    setImpactCategoryFilter('all');
    setImpactTypeFilter('all');
    setInstrumentWorkFilter('all');
    setInstrumentResourceFilter('all');
    setInstrumentFunctionFilter('all');
    setInstrumentStatusFilter('all');
    setInstrumentLocationFilter('all');
    setSelectedInstrumentAllocationId('');
    setProjects([]);
    setIssues([]);
    setIssueRelationshipsById({});
    setPublishedDocuments([]);
    setPublishedFolder(null);
    setDocumentListRows([]);
    setDocumentListSpreadsheet(null);
    setPublishedDocumentsPartial(false);
    setPublishedDocumentsMessage('');
    resetEapImport();
    setIssueTypeOptions([]);
    setCustomFieldDefinitions([]);
    setProjectUsers([]);
    setError('');

    if (!hubId) return;

    setProjectsLoading(true);

    try {
      const projectsData = await requestJson(`/api/hubs/${encodeURIComponent(hubId)}/projects`);
      setProjects(projectsData);
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setProjectsLoading(false);
    }
  }

  async function handleProjectSelect(projectId, options = {}) {
    setSelectedProjectId(projectId);
    setSelectedIssueId('');
    if (!options.preserveModule) setActiveModule('');
    setProjectPickerOpen(false);
    setIssueFilter('all');
    setCategoryFilter('all');
    setTypeFilter('all');
    setImpactCategoryFilter('all');
    setImpactTypeFilter('all');
    setInstrumentWorkFilter('all');
    setInstrumentResourceFilter('all');
    setInstrumentFunctionFilter('all');
    setInstrumentStatusFilter('all');
    setInstrumentLocationFilter('all');
    setInstrumentResourceTypeFilter('all');
    setSelectedInstrumentAllocationId('');
    setIssues([]);
    setIssueRelationshipsById({});
    setPublishedDocuments([]);
    setPublishedFolder(null);
    setDocumentListRows([]);
    setDocumentListSpreadsheet(null);
    setPublishedDocumentsPartial(false);
    setPublishedDocumentsMessage('');
    resetEapImport();
    setIssueTypeOptions([]);
    setCustomFieldDefinitions([]);
    setProjectUsers([]);
    setError('');

    if (!projectId) return;

    setIssuesLoading(true);

    try {
      let issuesWarning = '';
      const [issuesData, issueTypesData, fieldDefinitionsData, projectUsersData] = await Promise.all([
        requestJson(`/api/projects/${encodeURIComponent(projectId)}/issues`).catch((requestError) => {
          const message = String(requestError?.message || '');
          const missingResource = requestError?.status === 404 || normalizeText(message).includes('requested resource does not exist');

          if (missingResource) {
            issuesWarning = 'Este projeto nao possui o recurso de Issues habilitado no Autodesk Construction Cloud. Os modulos dependentes de Issues ficarao vazios ate selecionar outro projeto.';
            return [];
          }

          throw requestError;
        }),
        requestJson(`/api/projects/${encodeURIComponent(projectId)}/issue-types`).catch(() => []),
        requestJson(`/api/projects/${encodeURIComponent(projectId)}/issue-attribute-definitions`).catch(() => []),
        requestJson(`/api/projects/${encodeURIComponent(projectId)}/users`).catch(() => [])
      ]);
      const interfaceCategory = issueTypesData.find(
        (item) =>
          item.kind === 'type' &&
          normalizeText([item.title, item.category].filter(Boolean).join(' ')).includes(
            normalizeText('Interface e Coordenação Multidisciplinar')
          )
      );

      const loadedIssues = Array.isArray(issuesData) ? issuesData : issuesData.issues || [];
      const relationshipData = await requestJson(`/api/projects/${encodeURIComponent(projectId)}/issues/relationships`, {
        method: 'POST',
        body: JSON.stringify({ issueIds: loadedIssues.map((issue) => issue.id) })
      }).catch(() => ({ results: {} }));
      const relationshipMap = relationshipData.results || {};
      setIssueRelationshipsById(relationshipMap);
      const issuesWithRelationships = loadedIssues.map((issue) => {
        const refs = new Set();
        extractReferencedIssuesFromIssueDetail(issue).forEach((reference) => {
          if (reference.targetIssueId) refs.add(String(reference.targetIssueId));
          if (reference.targetDisplayId) refs.add(String(reference.targetDisplayId));
        });
        collectReferenceIds(relationshipMap[issue.id] || [], refs);
        return { ...issue, relationshipReferenceIds: [...refs] };
      });
      setIssues(issuesWithRelationships);
      setIssueTypeOptions(issueTypesData);
      setCustomFieldDefinitions(fieldDefinitionsData);
      setProjectUsers(projectUsersData);

      if (issuesWarning) {
        setError(issuesWarning);
      }
      setNewInterfaceIssue((issue) => ({
        ...issue,
        issueTypeId: interfaceCategory?.id || issue.issueTypeId,
        issueSubtypeId: interfaceCategory?.id === issue.issueTypeId ? issue.issueSubtypeId : ''
      }));
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setIssuesLoading(false);
    }
  }

  async function loadPublishedDocuments(projectId = selectedProjectId, options = {}) {
    if (!selectedHubId || !projectId) return;

    setDocumentsLoading(true);
    setDocumentsSpreadsheetSaveMessage('');

    try {
      const refreshParam = options.forceRefresh ? '&refresh=soft' : '';
      const documentsData = await requestJson(
        `/api/hubs/${encodeURIComponent(selectedHubId)}/projects/${encodeURIComponent(projectId)}/published-documents?limit=100${refreshParam}`
      );
      setPublishedFolder(documentsData.folder || null);
      setPublishedDocuments(documentsData.documents || []);
      setDocumentListRows(documentsData.rows || []);
      setDocumentListSpreadsheet(documentsData.spreadsheet || null);
      setPublishedDocumentsPartial(Boolean(documentsData.partial));
      setPublishedDocumentsMessage(documentsData.message || '');
      setDocumentsUpdatedAt(new Date().toISOString());
    } catch (requestError) {
      setPublishedFolder(null);
      setPublishedDocuments([]);
      setDocumentListRows([]);
      setDocumentListSpreadsheet(null);
      setPublishedDocumentsPartial(false);
      setPublishedDocumentsMessage(requestError.message || 'Nao foi possivel consultar a Autodesk neste momento.');
      setDocumentsUpdatedAt('');
      setError(requestError.message);
    } finally {
      setDocumentsLoading(false);
    }
  }


  async function updatePublishedDocumentsSpreadsheet() {
    if (!selectedHubId || !selectedProjectId) return;

    if (!documentListRows.length) {
      setDocumentsSpreadsheetSaveMessage('Carregue primeiro o comparativo de Documentos Emitidos antes de atualizar a planilha no ACC.');
      return;
    }

    const confirmed = window.confirm(
      'Esta acao ira atualizar a aba Central G5 da planilha no ACC e criar uma nova versao do arquivo. Deseja continuar?'
    );
    if (!confirmed) return;

    setDocumentsSpreadsheetSaving(true);
    setDocumentsSpreadsheetSaveMessage('Atualizando a planilha no ACC...');
    setError('');

    try {
      const result = await requestJson(
        `/api/hubs/${encodeURIComponent(selectedHubId)}/projects/${encodeURIComponent(selectedProjectId)}/published-documents/update-spreadsheet`,
        {
          method: 'POST',
          body: JSON.stringify({ limit: 100 })
        }
      );

      setDocumentsSpreadsheetSaveMessage(result.message || 'Planilha atualizada no ACC.');
      setDocumentListSpreadsheet((currentSpreadsheet) => ({
        ...(currentSpreadsheet || {}),
        ...(result.spreadsheet || {}),
        updatedAt: result.spreadsheet?.updatedAt || new Date().toISOString()
      }));
      setDocumentsUpdatedAt(new Date().toISOString());
    } catch (requestError) {
      const message = requestError.message || 'Nao foi possivel atualizar a planilha no ACC.';
      setDocumentsSpreadsheetSaveMessage(message);
      setError(message);
    } finally {
      setDocumentsSpreadsheetSaving(false);
    }
  }

  function resetEapImport() {
    setEapFileName('');
    setEapPreview(null);
    setEapResults([]);
    setEapDryRun(true);
    setEapExecutionHistory([]);
    setEapEditedRows([]);
    setEapEditedCells({});
  }

  function validateEapRow(row) {
    const errors = [];
    if (!row.title) errors.push('Linha sem Title.');
    if (!row.category) errors.push('Linha sem Category.');
    if (!row.issueType) errors.push('Linha sem Type.');
    if (!row.status) errors.push('Linha sem Status.');
    return { ...row, errors, validation: errors.length ? 'erro' : row.duplicate ? 'duplicado' : 'pronto' };
  }

  function updateEapCell(rowId, field, value) {
    setEapEditedRows((rows) =>
      rows.map((row) => {
        if (row.id !== rowId) return row;
        const payload = { ...(row.payload || {}) };
        if (field === 'title') payload.title = value;
        if (field === 'description') payload.description = value;
        if (field === 'status') payload.status = value;
        return validateEapRow({ ...row, [field]: value, payload });
      })
    );
    setEapEditedCells((cells) => ({ ...cells, [`${rowId}:${field}`]: true }));
  }

  function updateEapSourceCell(rowId, columnKey, value) {
    setEapEditedRows((rows) =>
      rows.map((row) => {
        if (row.id !== rowId) return row;
        const sourceValues = { ...(row.sourceValues || {}), [columnKey]: value };
        const normalized = normalizeText(columnKey.replace(/__\d+$/, ''));
        const nextRow = { ...row, sourceValues };
        if (normalized.includes('title')) nextRow.title = value;
        if (normalized.includes('description')) nextRow.description = value;
        if (normalized.includes('category')) nextRow.category = value;
        if (normalized === 'type' || normalized.includes(' type')) nextRow.issueType = value;
        if (normalized.includes('status')) nextRow.status = value;
        if (normalized.includes('assigned to')) nextRow.assignee = value;
        if (normalized.includes('codigo eap') || normalized.includes('código eap')) nextRow.eapCode = value;
        if (normalized.includes('due date')) nextRow.dueDate = value;
        if (nextRow.payload) {
          nextRow.payload = { ...nextRow.payload, title: nextRow.title, description: nextRow.description, status: nextRow.status };
        }
        return validateEapRow(nextRow);
      })
    );
    setEapEditedCells((cells) => ({ ...cells, [`${rowId}:${columnKey}`]: true }));
  }

  function resetImModule() {
    setImFileName('');
    setImPreview(null);
    setImResults([]);
    setImApplySummary(null);
    setImActionFeedback(null);
    setImAllowClearEmpty(false);
    setImSheetRows([]);
    setImSheetOriginalRows([]);
    setImSheetCustomColumns([]);
    setImSheetSelectedRows({});
    setImSheetChanges({});
    setImSheetPreviewRows([]);
  }

  async function loadImWorksheet() {
    if (!selectedProjectId) return setError('Selecione um projeto antes de carregar issues do ACC.');
    setImLoading(true);
    setError('');
    try {
      const [issuesData, customDefinitions] = await Promise.all([
        requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues?limit=200`),
        requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issue-attribute-definitions`).catch(() => [])
      ]);
      const customCols = (customDefinitions || []).map((field) => ({
        id: field.id,
        name: cleanCustomFieldName(field.name || field.title || field.id),
        type: normalizeCustomFieldType(field.type),
        required: Boolean(field.required),
        options: Array.isArray(field.options) ? field.options.map((option) => ({ value: String(option.value ?? option.id ?? option.name ?? option), label: String(option.label ?? option.name ?? option.value ?? option.id ?? option) })) : []
      }));
      const normalizedIssues = Array.isArray(issuesData) ? issuesData : issuesData.issues || [];
      const rows = normalizedIssues.map((issue) => {
        const customById = Object.fromEntries((issue.customAttributes || []).map((attr) => [attr.id, formatDetailValue(attr.rawValue ?? attr.value)]));
        const humanIssueId = String(issue.displayId ?? '').trim() || String(issue.id ?? '').trim();
        const row = {
          issueId: humanIssueId,
          accIssueId: issue.id,
          code: humanIssueId,
          title: issue.title || '',
          description: issue.description || '',
          status: issue.status || '',
          assignedTo: issue.assignedTo || '',
          dueDate: getDateInputValue(issue.dueDate),
          issueType: issue.issueType || '',
          issueSubtype: issue.issueSubtype || '',
          category: issue.category || '',
          priority: issue.priority || '',
          link: issue.webUrl || ''
        };
        customCols.forEach((col) => { row[`cf_${col.id}`] = customById[col.id] || ''; });
        return row;
      });
      setImSheetCustomColumns(customCols);
      setImSheetRows(rows);
      setImSheetOriginalRows(JSON.parse(JSON.stringify(rows)));
      setImSheetChanges({});
      setImSheetPreviewRows([]);
      console.info('[IM Worksheet] issues carregados', { totalIssues: rows.length, totalCustomFields: customCols.length });
    } catch (e) {
      setError(e.message);
    } finally {
      setImLoading(false);
    }
  }

  function updateWorksheetCell(rowIndex, field, value) {
    setImSheetRows((current) => current.map((row, idx) => (idx === rowIndex ? { ...row, [field]: value } : row)));
    setImSheetChanges((current) => ({ ...current, [`${rowIndex}:${field}`]: true }));
  }

  function downloadImReport() {
    if (!imPreview) return;
    const workbook = XLSX.utils.book_new();
    const summary = [
      { Métrica: 'Arquivo', Valor: imFileName || '-' },
      { Métrica: 'Aba', Valor: imPreview.sheetName || 'Config. issue' },
      { Métrica: 'Categorias reconhecidas', Valor: imPreview.summary?.categories || 0 },
      { Métrica: 'Tipos reconhecidos', Valor: imPreview.summary?.types || 0 },
      { Métrica: 'Campos reconhecidos', Valor: imPreview.summary?.fields || 0 },
      { Métrica: 'Matriz reconhecida', Valor: imPreview.summary?.matrix || 0 },
      { Métrica: 'Erros', Valor: imPreview.summary?.errors || 0 },
      { Métrica: 'Alertas', Valor: imPreview.summary?.warnings || 0 },
      { Métrica: 'Ignorados', Valor: imPreview.summary?.ignored || 0 },
      { Métrica: 'Execução', Valor: imResults.length ? 'Executada' : 'Somente validação' }
    ];
    XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(summary), 'Resumo');
    const addSheet = (name, rows) => {
      if (!rows?.length) return;
      XLSX.utils.book_append_sheet(workbook, XLSX.utils.json_to_sheet(rows), name);
    };
    addSheet('Categorias', imPreview.categories || []);
    addSheet('Tipos', imPreview.types || []);
    addSheet('Campos', imPreview.fields || []);
    addSheet('Matriz', imPreview.matrix || []);
    addSheet('ErrosAlertas', (imPreview.issues || []).filter((row) => row.status !== 'ok'));
    addSheet('Execucao', imResults || []);
    XLSX.writeFile(workbook, `relatorio-config-issues-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  async function handleImFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;
    if (!selectedProjectId) return setError('Selecione um projeto antes de carregar a planilha.');
    if (!file.name.toLowerCase().endsWith('.xlsx')) return setError('Selecione um arquivo Excel no formato .xlsx.');

    setError('');
    setImLoading(true);
    setImResults([]);
    setImPreview(null);
    setImFileName(file.name);
    if (imMode === 'create') setImPreview({ message: 'Planilha carregada com sucesso. Verificando dados...' });

    try {
      const endpoint = imMode === 'update' ? 'preview-update-file' : 'preview-create-file';
      const response = await fetch(`/api/projects/${encodeURIComponent(selectedProjectId)}/im-issues/${endpoint}`, {
        method: 'POST',
        credentials: 'include',
        headers: { 'Content-Type': 'application/octet-stream', 'X-File-Name': encodeURIComponent(file.name) },
        body: file
      });
      const payload = await response.json();
      if (!response.ok) throw new Error(payload?.message || 'Falha ao ler planilha IM.');
      setImPreview(payload);
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setImLoading(false);
    }
  }

  async function executeImFlow() {
    console.info('[IM UI] Clique em aplicar configuração', {
      mode: imMode,
      hasPreview: Boolean(imPreview),
      categories: imPreview?.categories?.length || 0,
      types: imPreview?.types?.length || 0,
      fields: imPreview?.fields?.length || 0,
      rows: imPreview?.rows?.length || 0,
      hasWorkbookData: Boolean(imPreview?.workbookData),
      selectedProjectId: selectedProjectId || ''
    });
    if (!selectedProjectId) {
      const message = 'Selecione um projeto antes de aplicar a configuração no ACC.';
      setError(message);
      setImActionFeedback({ type: 'error', message });
      console.warn('[IM UI] Fluxo interrompido: projeto não selecionado');
      return;
    }
    if (imMode === 'update' && !imPreview?.rows?.length) {
      const message = 'Não há linhas elegíveis para atualização. Revise a planilha.';
      setImActionFeedback({ type: 'warning', message });
      console.warn('[IM UI] Fluxo interrompido: sem linhas para atualizar');
      return;
    }
    if (imMode === 'update') {
      const total = imPreview?.rows?.length || 0;
      const ready = imPreview?.readyRows || 0;
      const errors = (imPreview?.rows || []).filter((row) => row.validation === 'erro').length;
      const ignored = (imPreview?.rows || []).filter((row) => row.validation === 'ignorado').length;
      const confirmed = window.confirm(
        `Confirma atualização no ACC?\nLinhas lidas: ${total}\nProntas: ${ready}\nCom erro: ${errors}\nIgnoradas: ${ignored}\n\nEsta ação altera dados diretamente no Autodesk Construction Cloud e atualizará a planilha com o resultado por linha.`
      );
      if (!confirmed) return;
    }
    if (imMode === 'create' && !imPreview?.categories?.length && !imPreview?.types?.length && !imPreview?.fields?.length) {
      const message = 'Prévia sem itens para criar. Verifique a aba Config. issue antes de aplicar.';
      setImActionFeedback({ type: 'warning', message });
      console.warn('[IM UI] Fluxo interrompido: sem itens para criar');
      return;
    }
    if (!imPreview?.workbookData) {
      const message = 'A prévia não contém workbookData. Recarregue a planilha antes de aplicar.';
      setImActionFeedback({ type: 'error', message });
      setError(message);
      console.warn('[IM UI] Fluxo interrompido: workbookData ausente');
      return;
    }
    setImLoading(true);
    setError('');
    setImActionFeedback({ type: 'info', message: 'Aplicando configuração no ACC... aguarde.' });
    try {
      const endpoint = imMode === 'update' ? 'update' : 'create';
      console.info('[IM UI] Disparando requisição final', {
        endpoint: `/api/projects/${encodeURIComponent(selectedProjectId)}/im-issues/${endpoint}`,
        method: 'POST'
      });
      const result = await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/im-issues/${endpoint}`, {
        method: 'POST',
        body: JSON.stringify({ workbookData: imPreview.workbookData, allowClearEmpty: imAllowClearEmpty })
      });
      console.info('[IM UI] Resposta recebida da aplicação IM', {
        totalRows: result?.totalRows || 0,
        createdRows: result?.createdRows || 0,
        errorRows: result?.errorRows || 0,
        persistenceVerified: Boolean(result?.persistenceVerified)
      });
      setImResults(result.results || []);
      setImApplySummary({
        message: result.message || '',
        persistenceVerified: Boolean(result.persistenceVerified),
        persistedChecks: result.persistedChecks || {},
        errorRows: result.errorRows || 0
      });
      setImActionFeedback({
        type: result.persistenceVerified ? 'success' : 'warning',
        message: result.persistenceVerified
          ? 'Configuração aplicada e persistência confirmada no ACC.'
          : 'Processamento concluído, mas nem todos os itens foram confirmados no ACC. Revise a tabela de resultados.'
      });
      await handleProjectSelect(selectedProjectId, { preserveModule: true });
    } catch (requestError) {
      const authExpired = Number(requestError?.status) === 401 || String(requestError?.message || '').includes('entrar com sua conta Autodesk');
      const feedbackMessage = authExpired
        ? 'Sua sessão Autodesk expirou ou não está ativa. Clique em "Entrar com Autodesk" para autenticar novamente e reaplique a configuração.'
        : `Falha ao aplicar configuração: ${requestError.message}`;
      if (authExpired) setUser(null);
      setError(feedbackMessage);
      setImActionFeedback({ type: 'error', message: feedbackMessage });
      console.error('[IM UI] Falha no fluxo final de aplicação', requestError);
    } finally {
      setImLoading(false);
    }
  }

  async function handleEapFileChange(event) {
    const file = event.target.files?.[0];
    if (!file) return;

    setError('');
    setEapLoading(true);
    setEapResults([]);
    setEapPreview(null);
    setEapFileName(file.name);

    try {
      if (!file.name.toLowerCase().endsWith('.xlsx')) {
        throw new Error('Selecione um arquivo Excel no formato .xlsx.');
      }

      if (file.size > 25 * 1024 * 1024) {
        throw new Error('Esta planilha esta muito grande para importar aqui. Envie uma copia menor ou sem abas auxiliares pesadas.');
      }

      const response = await fetch(`/api/projects/${encodeURIComponent(selectedProjectId)}/eap-issues/preview-file`, {
        method: 'POST',
        credentials: 'include',
        headers: {
          'Content-Type': 'application/octet-stream',
          'X-File-Name': encodeURIComponent(file.name)
        },
        body: file
      });
      const responseText = await response.text();
      let preview = {};

      try {
        preview = responseText ? JSON.parse(responseText) : {};
      } catch {
        throw new Error('O servidor retornou uma resposta inesperada ao ler a planilha ACS Build.');
      }

      if (!response.ok) {
        throw new Error(preview.message || preview.error || responseText || 'Nao foi possivel ler a planilha ACS Build.');
      }

      setEapPreview(preview);
      setEapEditedRows((preview.rows || []).map((row) => ({ ...row })));
      setEapEditedCells({});
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setEapLoading(false);
    }
  }

  async function createEapIssues() {
    if (!eapPreview?.rows?.length) return;

    setError('');
    setEapLoading(true);

    try {
      const result = await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/eap-issues/create`, {
        method: 'POST',
        body: JSON.stringify({
          dryRun: eapDryRun,
          rows: eapEditedRows.length ? eapEditedRows : eapPreview.rows,
          workbookData: eapPreview.workbookData
        })
      });
      setEapResults(result.results || []);
      setEapExecutionHistory((currentHistory) => [
        {
          id: `${Date.now()}-${Math.random().toString(16).slice(2, 8)}`,
          createdAt: new Date().toISOString(),
          dryRun: Boolean(result.dryRun),
          total: (result.results || []).length,
          created: (result.results || []).filter((row) => row.status === 'criado').length,
          duplicated: (result.results || []).filter((row) => row.status === 'duplicado').length,
          errors: (result.results || []).filter((row) => row.status === 'erro').length
        },
        ...currentHistory
      ].slice(0, 15));
      if (!result.dryRun) {
        await handleProjectSelect(selectedProjectId);
        setActiveModule('eap-import');
      }
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setEapLoading(false);
    }
  }

  function exportEapResultsXlsx() {
    const rows = (eapResults.length ? eapResults : eapPreview?.rows || []).map((row) => ({
      Linha: row.line,
      'Codigo EAP': row.eapCode || '',
      'Nome do Issue': row.title || '',
      'Issue criado?': row.status === 'criado' ? 'Sim' : row.validation === 'duplicado' || row.status === 'duplicado' ? 'Duplicado' : 'Nao',
      'ID do Issue ACC': row.issueId || '',
      'Link do Issue ACC': row.issueLink || '',
      'Data/hora da criacao': row.createdAt ? formatDateTime(row.createdAt) : '',
      'Campos nativos preenchidos': (row.nativeFieldsFilled || []).join(', '),
      'Campos personalizados preenchidos': (row.customFieldsFilled || []).join(', '),
      'Campos nao preenchidos': (row.customFieldsMissing || []).join(', '),
      'Mensagem de erro': row.message || row.errors?.join(' ') || row.warnings?.join(' ') || ''
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Retorno Issues ACS Build');
    XLSX.writeFile(workbook, `retorno-issues-acs-build-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  const disciplineColors = {
    'todas as disciplinas': '#8b95a1',
    geotecnia: '#13a89e',
    geologia: '#6f4cc3',
    estruturas: '#2f80ed',
    'recursos hidricos': '#1f9bb4',
    hidraulica: '#1f9bb4',
    eletrica: '#f2c94c',
    mecanica: '#f2994a',
    producao: '#eb5757',
    planejamento: '#56cc9d',
    gestao: '#0b7285',
    bim: '#9b51e0',
    comercial: '#27ae60',
    financeiro: '#bb6bd9',
    rh: '#ff7aa2',
    ssma: '#2d9cdb',
    'meio ambiente': '#98a824',
    'gestao da informacao': '#0b7285'
  };

  function getDisciplineColor(discipline) {
    const key = normalizeText(discipline || '');
    const hit = Object.entries(disciplineColors).find(([name]) => key.includes(name));
    return hit?.[1] || '#ea5455';
  }

function exportIssueLinksManagementSheet() {
    const rows = filteredLinkFlows.map((issue) => ({
      'ID do issue ACC': issue.id,
      Título: issue.title || '',
      Descrição: issue.description || '',
      Categoria: issue.category || '',
      Tipo: issue.issueType || '',
      'Tipo (ID)': issue.raw?.issueTypeId || issue.raw?.typeId || '',
      Status: issue.status || '',
      Responsável: issue.assignedTo || '',
      'Data de vencimento': issue.effectiveDate || issue.dueDate || '',
      'Data de criação': issue.createdAt || '',
      'Data de atualização': issue.updatedAt || '',
      Disciplina: getIssueDiscipline(issue) || '',
      'Código EAP': getIssueCustomValue(issue, ['codigo eap', 'código eap']),
      'Pacote de Trabalho': getIssueCustomValue(issue, ['pacote de trabalho']),
      'Marco Contratual': getIssueCustomValue(issue, ['marco contratual']),
      'Impacto no Prazo': getIssueCustomValue(issue, ['impacto no prazo', 'impacto no cronograma']),
      'Impacto no Escopo': getIssueCustomValue(issue, ['impacto no escopo']),
      'Impacto na Medição': getIssueCustomValue(issue, ['impacto na medicao', 'impacto em medicao']),
      'Status da Medição': getIssueCustomValue(issue, ['status da medicao']),
      'Pacote/ Marco contratual': getCustomFieldValue(issue, customFieldDefinitions, 'Pacote/ Marco contratual') || '',
      Publicada: issue.published === false ? 'Não' : 'Sim',
      'Tipo de linha': 'Issue secundário',
      Aberto: isOpenIssue(issue) ? 'Sim' : 'Não',
      Fechado: isOpenIssue(issue) ? 'Não' : 'Sim',
      Atrasado: isOverdue(issue) ? 'Sim' : 'Não',
      'Link ACC': issue.raw?.links?.webView?.href || ''
    }));
    const worksheet = XLSX.utils.json_to_sheet(rows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Gantt_EAP');
    XLSX.writeFile(workbook, `gantt-eap-issues-${new Date().toISOString().slice(0, 10)}.xlsx`);
  }

  function openDocumentsModule() {
    setActiveModule('documents');
    if (documentListRows.length === 0 && publishedDocuments.length === 0 && !documentsLoading) {
      loadPublishedDocuments();
    }
  }

  function changeDocumentSort(field) {
    setDocumentSort((currentSort) => ({
      field,
      direction: currentSort.field === field && currentSort.direction === 'asc' ? 'desc' : 'asc'
    }));
  }

  function syncDocumentTableScroll(source) {
    const topScroller = documentTopScrollRef.current;
    const tableScroller = documentTableScrollRef.current;
    if (!topScroller || !tableScroller) return;

    if (source === 'top') {
      tableScroller.scrollLeft = topScroller.scrollLeft;
    } else {
      topScroller.scrollLeft = tableScroller.scrollLeft;
    }
  }

  function clearDocumentFilters() {
    setDocumentSearch('');
    setDocumentStatusFilter('all');
    setDocumentTypeFilter('all');
  }

  async function refreshCurrentModule() {
    if (!selectedProjectId) return;
    if (activeModule === 'documents') {
      await loadPublishedDocuments();
      return;
    }
    await handleProjectSelect(selectedProjectId, { preserveModule: true });
  }

  function selectMeetingTemplate(templateId) {
    const template = meetingTemplates.find((item) => item.id === templateId) || meetingTemplates[0];
    setSelectedMeetingTemplateId(template.id);
    setMeetingCustomTopics(template.topics.join('\n'));
  }

  function exportMeetingDraftWord() {
    const template = meetingTemplates.find((item) => item.id === selectedMeetingTemplateId) || meetingTemplates[0];
    const topics = meetingCustomTopics
      .split('\n')
      .map((item) => item.trim())
      .filter(Boolean);
    const agendaRows = topics.map((topic, index) => `
      <tr>
        <td>${index + 1}</td>
        <td>${htmlEscape(topic)}</td>
        <td></td>
        <td></td>
      </tr>
    `).join('');
    const html = `<!doctype html>
<html>
  <head>
    <meta charset="utf-8">
    <title>${htmlEscape(template.title)}</title>
    <style>
      body { font-family: Arial, sans-serif; color: #163a3d; line-height: 1.45; }
      h1 { color: #0f5f69; font-size: 26px; margin-bottom: 4px; }
      h2 { color: #177f87; font-size: 16px; margin-top: 28px; text-transform: uppercase; }
      .meta { border: 1px solid #b9d6d5; border-collapse: collapse; width: 100%; margin: 18px 0 24px; }
      .meta td { border: 1px solid #b9d6d5; padding: 8px 10px; vertical-align: top; }
      .label { width: 160px; background: #eef7f6; font-weight: bold; }
      table { border-collapse: collapse; width: 100%; }
      th { background: #0f5f69; color: #fff; text-align: left; padding: 8px; }
      td { border: 1px solid #cfe2e1; padding: 8px; min-height: 26px; }
      .small { color: #58706f; font-size: 12px; margin-top: 28px; }
    </style>
  </head>
  <body>
    <h1>${htmlEscape(template.title)}</h1>
    <p>${htmlEscape(template.purpose)}</p>
    <table class="meta">
      <tr><td class="label">Projeto</td><td>${htmlEscape(selectedProject?.name || selectedProjectId || 'A definir')}</td></tr>
      <tr><td class="label">Data</td><td>${htmlEscape(meetingDate ? formatDate(meetingDate) : 'A definir')}</td></tr>
      <tr><td class="label">Participantes</td><td>${htmlEscape(meetingParticipants || 'A definir')}</td></tr>
    </table>
    <h2>Pauta</h2>
    <table>
      <thead>
        <tr><th style="width: 48px;">Item</th><th>Ponto de pauta</th><th>Encaminhamento</th><th style="width: 180px;">Responsavel / Prazo</th></tr>
      </thead>
      <tbody>
        ${agendaRows || '<tr><td>1</td><td>Pauta a definir</td><td></td><td></td></tr>'}
      </tbody>
    </table>
    <h2>Decisoes e proximos passos</h2>
    <table>
      <thead>
        <tr><th>Decisao / acao</th><th style="width: 180px;">Responsavel</th><th style="width: 140px;">Prazo</th><th style="width: 140px;">Status</th></tr>
      </thead>
      <tbody>
        <tr><td>&nbsp;</td><td></td><td></td><td></td></tr>
        <tr><td>&nbsp;</td><td></td><td></td><td></td></tr>
        <tr><td>&nbsp;</td><td></td><td></td><td></td></tr>
      </tbody>
    </table>
    <p class="small">Documento gerado no app Central G5 para apoiar registro de reunioes e posterior preenchimento no ACC.</p>
  </body>
</html>`;
    const blob = new Blob([`\uFEFF${html}`], { type: 'application/msword;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `roteiro-reuniao-${template.id}-${new Date().toISOString().slice(0, 10)}.doc`;
    link.click();
    URL.revokeObjectURL(url);
  }

  function exportPublishedDocumentsCsv() {
    const headers = ['Disciplina', 'Codigo de engenharia', 'Descricao', 'Tipo de documento', 'Nome', 'Rev.', 'Versao', 'GRD', 'Status', 'Data de emissao'];
    const rows = sortedDocumentRows.map((document) => [
      document.discipline || '',
      document.code || '',
      document.title || '',
      document.documentType || '',
      document.emittedFileName || '',
      document.emittedRevision || '',
      document.emittedVersion || '',
      document.emittedGrd || '',
      document.status || (document.emitted ? 'Emitido' : 'Pendente'),
      document.emittedDate || document.emissionDate || ''
    ]);
    const csv = [headers, ...rows].map((row) => row.map(csvEscape).join(';')).join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `documentos-emitidos-${selectedProject?.name || selectedProjectId}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  function exportPublishedDocumentsPowerBiCsv() {
    exportPublishedDocumentsCsv();
  }

  function handleBimParameterDraftChange(field, value) {
    setBimParameterDraft((current) => ({ ...current, [field]: value }));
  }

  function addBimParameterPlanItem(event) {
    event.preventDefault();
    const parameterName = bimParameterDraft.name.trim();

    if (!parameterName) {
      setBimParameterMessage('Informe o nome do parâmetro antes de adicionar ao plano.');
      return;
    }

    const item = {
      id: `${Date.now()}-${bimParameterPlan.length + 1}`,
      name: parameterName,
      type: bimParameterDraft.type,
      target: bimParameterDraft.target,
      value: bimParameterDraft.value.trim(),
      notes: bimParameterDraft.notes.trim(),
      status: 'Pronto para validar no ACC'
    };

    setBimParameterPlan((current) => [...current, item]);
    setBimParameterDraft((current) => ({
      name: '',
      type: current.type,
      target: current.target,
      value: '',
      notes: ''
    }));
    setBimParameterMessage('Parâmetro adicionado ao plano de coordenação BIM.');
  }

  function exportBimParameterPlanCsv() {
    if (!bimParameterPlan.length) {
      setBimParameterMessage('Adicione pelo menos um parâmetro ao plano antes de exportar.');
      return;
    }

    const headers = ['Parametro', 'Tipo', 'Aplicacao', 'Valor padrao', 'Observacoes', 'Status'];
    const rows = bimParameterPlan.map((item) => [
      item.name,
      item.type,
      item.target,
      item.value,
      item.notes,
      item.status
    ]);
    const csv = [headers, ...rows].map((row) => row.map(csvEscape).join(';')).join('\n');
    const blob = new Blob([`\uFEFF${csv}`], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `plano-parametros-bim-${selectedProject?.name || selectedProjectId || 'projeto'}.csv`;
    link.click();
    URL.revokeObjectURL(url);
  }

  async function updateIssue(issueId, patch) {
    setSavingIssueId(issueId);
    setError('');

    try {
      const updatedIssue = await requestJson(
        `/api/projects/${encodeURIComponent(selectedProjectId)}/issues/${encodeURIComponent(issueId)}`,
        {
          method: 'PATCH',
          body: JSON.stringify(patch)
        }
      );

      setIssues((currentIssues) =>
        currentIssues.map((issue) => (issue.id === issueId ? { ...issue, ...updatedIssue } : issue))
      );
      setSelectedIssueDetails((currentIssue) =>
        currentIssue?.id === issueId ? { ...currentIssue, ...updatedIssue } : currentIssue
      );
    } catch (requestError) {
      setError(requestError.message);
    } finally {
      setSavingIssueId('');
    }
  }

  async function updateCustomAttribute(issue, field, value) {
    let foundAttribute = false;
    const nextAttributes = (issue.customAttributes || []).map((attribute) => {
      const attributeId = attribute.id || attribute.attributeDefinitionId;
      const sameField =
        String(attributeId || '') === String(field.id || field.attributeDefinitionId || '') ||
        normalizeFieldKey(attribute.name) === normalizeFieldKey(field.name);
      if (sameField) foundAttribute = true;
      return {
        attributeDefinitionId: attributeId,
        value: sameField ? value : attribute.rawValue
      };
    }).filter((attribute) => attribute.attributeDefinitionId);

    if (!foundAttribute && (field.id || field.attributeDefinitionId)) {
      nextAttributes.push({
        attributeDefinitionId: field.id || field.attributeDefinitionId,
        value
      });
    }

    await updateIssue(issue.id, { customAttributes: nextAttributes });
  }

  async function updateSingleCustomAttribute(issue, field, value) {
    if (!issue?.id || !(field?.id || field?.attributeDefinitionId)) return;
    const prepared = prepareCustomFieldWriteValue(field, value);
    if (!prepared.valid) {
      throw new Error(`O valor "${value}" nao existe nas opcoes do campo "${field.name || field.title || field.displayName || 'personalizado'}" neste projeto do ACC.`);
    }
    await updateIssue(issue.id, {
      customAttributes: [{
        attributeDefinitionId: field.id || field.attributeDefinitionId,
        value: prepared.value
      }]
    });
  }

  function findCustomFieldDefinitionByAliases(aliases) {
    const targets = aliases.map(normalizeFieldKey);
    return (customFieldDefinitions || []).find((field) => {
      const names = [field.id, field.definitionId, field.name, field.title, field.displayName, field.label]
        .filter(Boolean)
        .map(normalizeFieldKey);
      return names.some((name) => targets.includes(name) || targets.some((target) => name.includes(target)));
    });
  }

  function getCustomFieldOptions(field) {
    if (field?.options?.length) return field.options;
    const definition = (customFieldDefinitions || []).find((item) => {
      const fieldIds = [field?.id, field?.attributeDefinitionId].filter(Boolean).map(String);
      const definitionIds = [item.id, item.definitionId, item.attributeDefinitionId].filter(Boolean).map(String);
      const sameId = fieldIds.some((id) => definitionIds.includes(id));
      const sameName = normalizeFieldKey(field?.name) && normalizeFieldKey(field?.name) === normalizeFieldKey(item.name || item.title || item.displayName || item.label);
      return sameId || sameName;
    });
    return definition?.options || [];
  }

  async function updateProjectManagementMeasurementValue(row, rawValue) {
    const issue = issues.find((item) => item.id === row.measurementIssueId) || row.issues?.[0];
    const field = findCustomFieldDefinitionByAliases([
      'Valor da Medição',
      'Valor da Medicao',
      'Valor Medição',
      'Valor Medicao'
    ]);

    if (!issue?.id) {
      setError('Não encontrei uma issue segura para gravar o Valor da Medição deste marco.');
      return;
    }

    if (!field?.id) {
      setError('Não encontrei o campo personalizado "Valor da Medição" neste projeto do ACC.');
      return;
    }

    const nextValue = parseCurrencyNumber(rawValue);
    let foundAttribute = false;
    const nextAttributes = (issue.customAttributes || []).map((attribute) => {
      const sameField =
        String(attribute.id || attribute.attributeDefinitionId || '') === String(field.id) ||
        normalizeFieldKey(attribute.name) === normalizeFieldKey(field.name);
      if (sameField) foundAttribute = true;
      return {
        attributeDefinitionId: attribute.id || attribute.attributeDefinitionId,
        value: sameField ? nextValue : attribute.rawValue
      };
    }).filter((attribute) => attribute.attributeDefinitionId);

    if (!foundAttribute) {
      nextAttributes.push({
        attributeDefinitionId: field.id,
        value: nextValue
      });
    }

    await updateIssue(issue.id, { customAttributes: nextAttributes });
  }

  function resolveCustomFieldOptionValue(field, label) {
    const options = field?.options || field?.allowedValues || field?.values || [];
    const match = options.find((option) => {
      const optionText = option.label || option.name || option.title || option.value || option.id;
      return normalizeText(optionText) === normalizeText(label);
    });
    return match ? (match.id || match.value || match.label || label) : label;
  }

  function prepareCustomFieldWriteValue(field, value) {
    const options = getCustomFieldOptions(field);
    if (!options.length || value === null || value === undefined || value === '') {
      return { valid: true, value };
    }
    const normalizedValue = normalizeText(value);
    const match = options.find((option) => {
      const candidates = [
        option.id,
        option.value,
        option.label,
        option.name,
        option.title,
        option.displayName
      ].filter(Boolean);
      return candidates.some((candidate) => normalizeText(candidate) === normalizedValue);
    });
    if (!match) return { valid: false, value };
    return { valid: true, value: match.id || match.value || match.label || value };
  }

  async function updateProjectManagementMeasurementStatus(row, nextStatus) {
    const issue = issues.find((item) => item.id === row.measurementIssueId) || row.issues?.[0];
    const field = findCustomFieldDefinitionByAliases([
      'Status medição',
      'Status da Medição',
      'Status medição',
      'Status da Medição',
      'Status Medicao',
      'Status da Medicao',
      'Status Vale',
      'Medição',
      'Medição',
      'Medicao'
    ]);

    if (!issue?.id) {
      setError('Não encontrei uma issue segura para gravar o Status de medição deste marco.');
      return;
    }

    if (!field?.id) {
      setError('Não encontrei o campo personalizado "Status medição" neste projeto do ACC.');
      return;
    }

    const writeValue = resolveCustomFieldOptionValue(field, nextStatus);
    let foundAttribute = false;
    const nextAttributes = (issue.customAttributes || []).map((attribute) => {
      const sameField =
        String(attribute.id || attribute.attributeDefinitionId || '') === String(field.id) ||
        normalizeFieldKey(attribute.name) === normalizeFieldKey(field.name);
      if (sameField) foundAttribute = true;
      return {
        attributeDefinitionId: attribute.id || attribute.attributeDefinitionId,
        value: sameField ? writeValue : attribute.rawValue
      };
    }).filter((attribute) => attribute.attributeDefinitionId);

    if (!foundAttribute) {
      nextAttributes.push({
        attributeDefinitionId: field.id,
        value: writeValue
      });
    }

    await updateIssue(issue.id, { customAttributes: nextAttributes });
  }

  function updateNewCustomAttribute(fieldId, value) {
    setNewInterfaceIssue((issue) => ({
      ...issue,
      customAttributes: {
        ...issue.customAttributes,
        [fieldId]: value
      }
    }));
  }

  async function createInterfaceIssue(event) {
    event.preventDefault();
    setError('');

    try {
      await requestJson('/api/auth/keep-alive', { method: 'POST' });

      const customAttributes = Object.entries(newInterfaceIssue.customAttributes || {})
        .filter(([, value]) => value !== undefined && value !== null && value !== '')
        .map(([attributeDefinitionId, value]) => ({
          attributeDefinitionId,
          value
        }));

      await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues`, {
        method: 'POST',
        body: JSON.stringify({
          title: newInterfaceIssue.title,
          dueDate: newInterfaceIssue.dueDate,
          description: newInterfaceIssue.description,
          issueTypeId: newInterfaceIssue.issueTypeId,
          issueSubtypeId: newInterfaceIssue.issueSubtypeId,
          assignedTo: newInterfaceIssue.assignedTo,
          customAttributes
        })
      });

      setNewInterfaceIssue((issue) => ({
        ...issue,
        title: '',
        dueDate: '',
        description: '',
        assignedTo: '',
        customAttributes: {}
      }));
      await handleProjectSelect(selectedProjectId, { preserveModule: true });
      setActiveModule('interface');
    } catch (requestError) {
      setError(requestError.message);
    }
  }

  async function logout() {
    await fetch('/api/auth/logout', {
      method: 'POST',
      credentials: 'include'
    });

    setUser(null);
    setHubs([]);
    setProjects([]);
    setIssues([]);
    setPublishedDocuments([]);
    setPublishedFolder(null);
    setDocumentListRows([]);
    setDocumentListSpreadsheet(null);
    setPublishedDocumentsPartial(false);
    setPublishedDocumentsMessage('');
    resetEapImport();
    setIssueTypeOptions([]);
    setCustomFieldDefinitions([]);
    setProjectUsers([]);
    setSelectedHubId('');
    setSelectedProjectId('');
    setSelectedIssueId('');
    setActiveModule('');
    setProjectPickerOpen(false);
    setIssueFilter('all');
    setCategoryFilter('all');
    setTypeFilter('all');
    setImpactCategoryFilter('all');
    setImpactTypeFilter('all');
    setError('');
  }

  return (
    <main className={`app-shell ${activeModule ? 'app-shell-operational' : ''}`}>
      <section className={`workspace ${activeModule ? 'workspace-operational' : 'workspace-home'}`}>
        <div className="topbar">
          <div className="brand-heading">
            <img src="/brand/g5-logo.png" alt="G5 Engenharia e Instrumentos" />
            <div>
              <p className="eyebrow">Etapa 2</p>
              <h1>Central G5</h1>
            </div>
          </div>
        </div>

        {user && (
          <div className="session-actions">
            <button className="ghost-button" type="button" onClick={logout}>
              Sair
            </button>
          </div>
        )}

        {!user && (
          <div className="login-panel">
            <h2>Conectar ao Autodesk Construction Cloud / Forma</h2>
            <p>Entre com sua conta Autodesk para acompanhar e controlar issues dos projetos.</p>
            <a className="primary-button" href="/api/auth/login">
              Entrar com Autodesk
            </a>
          </div>
        )}

        {user && !selectedBusinessUnit && (
          <section className="business-unit-panel">
            <p className="eyebrow">Unidade de Negocios</p>
            <h2>Escolha a area da G5</h2>
            <p>Esta escolha define quais projetos aparecem e qual interface sera carregada.</p>
            <div className="business-unit-grid">
              {businessUnitOptions.map((unit) => (
                <button key={unit.value} type="button" className="business-unit-card" onClick={() => handleBusinessUnitSelect(unit.value)}>
                  <span>{unit.title}</span>
                  <strong>{unit.title}</strong>
                  <p>{unit.description}</p>
                </button>
              ))}
            </div>
          </section>
        )}

        {user && selectedBusinessUnit && (!selectedProjectId || projectPickerOpen) && (
          <div className="content-grid">
            <aside className="account-panel">
              <p className="label">Usuário autenticado</p>
              <strong>{user.name}</strong>
              {user.email && <span>{user.email}</span>}
              <div className="business-unit-account">
                <span>Unidade selecionada</span>
                <strong>{selectedBusinessUnit}</strong>
                <button type="button" onClick={handleBusinessUnitReset}>Alterar unidade</button>
              </div>
            </aside>

            <section className="data-panel">
              <label htmlFor="hub-select">Hub</label>
              <select id="hub-select" value={selectedHubId} onChange={handleHubChange}>
                <option value="">Selecione um Hub</option>
                {hubs.map((hub) => (
                  <option key={hub.id} value={hub.id}>
                    {hub.name}
                  </option>
                ))}
              </select>

              <div className="list-header">
                <h2>Projetos</h2>
                {projectsLoading ? <span>Carregando...</span> : selectedHubId && <span>{filteredProjects.length} projetos encontrados</span>}
              </div>

              {!selectedHubId && <p className="muted">Escolha um Hub para listar os Projetos.</p>}
              {selectedHubId && !projectsLoading && filteredProjects.length === 0 && (
                <p className="muted">Nenhum projeto ativo foi encontrado neste Hub para o seu usuario.</p>
              )}

              {selectedProject && !projectPickerOpen && (
                <div className="selected-project-card">
                  <div>
                    <span>Projeto selecionado</span>
                    <strong>{selectedProject.name}</strong>
                  </div>
                  <button type="button" onClick={() => setProjectPickerOpen(true)}>
                    Trocar projeto
                  </button>
                </div>
              )}

              {(!selectedProjectId || projectPickerOpen) && (
                <ul className="project-list">
                  {filteredProjects
                    .slice()
                    .sort(sortByName)
                    .map((project) => (
                      <li key={project.id} className={project.id === selectedProjectId ? 'selected-item' : ''}>
                        <button type="button" onClick={() => handleProjectSelect(project.id)}>
                          <span>
                            <strong>{project.name}</strong>
                          </span>
                          <em>{project.id === selectedProjectId ? 'Selecionado' : 'Selecionar'}</em>
                        </button>
                      </li>
                    ))}
                </ul>
              )}
            </section>
          </div>
        )}

        {user && selectedProjectId && selectedBusinessUnit === 'G5 Instrumentos' && (
          <section className="module-panel instrumentos-shell">
            <div className="section-heading">
              <div>
                <p className="eyebrow">G5 Instrumentos</p>
                <h2>Instrumentacao</h2>
                <p className="selected-project-name">
                  Projeto selecionado: <strong>{selectedProject?.name || selectedProjectId}</strong>
                </p>
              </div>
              <div className="button-row">
                <button type="button" className="ghost-button" onClick={() => setProjectPickerOpen(true)}>Trocar projeto</button>
                <button type="button" className="ghost-button" onClick={handleBusinessUnitReset}>Alterar unidade</button>
                <button type="button" className="ghost-button" onClick={refreshCurrentModule}>Atualizar</button>
              </div>
            </div>

            <div className="instrumentos-layout">
              <aside className="instrumentos-sidebar">
                <p className="eyebrow">G5 Instrumentos</p>
                {instrumentSectionOptions.map((section) => (
                  <button
                    key={section.id}
                    className={`instrumentos-menu-item ${activeInstrumentSection === section.id ? 'active' : ''}`}
                    type="button"
                    onClick={() => setActiveInstrumentSection(section.id)}
                  >
                    <span className="module-icon" aria-hidden="true"><UiIcon name={section.icon} /></span>
                    <span>
                      <strong>{section.title}</strong>
                      <small>{section.subtitle}</small>
                    </span>
                  </button>
                ))}
              </aside>

              <section className="instrumentos-content">
                <div className="instrumentos-hero">
                  <div>
                    <p className="eyebrow">Controle Integrado de Obras</p>
                    <h3>{instrumentSectionOptions.find((section) => section.id === activeInstrumentSection)?.title || 'Dashboard Geral'}</h3>
                    <p>
                      Central gerencial para obras de instrumentacao, equipe, equipamentos, materiais, compras e pendencias usando os Issues visiveis do ACC.
                    </p>
                  </div>
                  <strong>{instrumentationIssues.length} issues encontrados</strong>
                </div>

                {activeInstrumentSection === 'dashboard' && (
                  <>
                    <div className="instrumentos-indicators instrumentos-indicators-large">
                      <div><span>Obras ativas</span><strong>{instrumentWorks.length}</strong></div>
                      <div><span>Colaboradores ativos</span><strong>{instrumentationIndicators.activeResources}</strong></div>
                      <div><span>Colaboradores alocados</span><strong>{instrumentationIndicators.allocations}</strong></div>
                      <div><span>Conflitos</span><strong>{instrumentationIndicators.conflicts}</strong></div>
                      <div><span>Equipamentos em obra</span><strong>{instrumentEquipmentItems.filter((item) => normalizeText(item.status).includes('obra')).length}</strong></div>
                      <div><span>Equipamentos cadastrados</span><strong>{instrumentEquipmentItems.length}</strong></div>
                      <div><span>Materiais pendentes</span><strong>{instrumentPurchaseStats.pending}</strong></div>
                      <div><span>Materiais recebidos</span><strong>{instrumentPurchaseStats.received}</strong></div>
                      <div><span>Issues em atraso</span><strong>{instrumentationIndicators.overdue}</strong></div>
                      <div><span>Sem responsavel</span><strong>{instrumentationIndicators.missingResponsible}</strong></div>
                      <div><span>Sem inicio</span><strong>{instrumentationIndicators.missingStart}</strong></div>
                      <div><span>Sem vencimento</span><strong>{instrumentationIndicators.missingEnd}</strong></div>
                    </div>

                    <div className="instrument-dashboard-grid">
                      <section className="instrument-management-card">
                        <p className="eyebrow">Alertas operacionais</p>
                        <h4>Pendencias criticas</h4>
                        <div className="instrument-alert-list">
                          {instrumentAlerts.slice(0, 8).map((alert, index) => (
                            <div key={`${alert.area}:${index}`} className={`instrument-alert-item level-${normalizeText(alert.level)}`}>
                              <span>{alert.level}</span>
                              <strong>{alert.area}</strong>
                              <p>{alert.text}</p>
                            </div>
                          ))}
                          {!instrumentAlerts.length && <p className="muted-text">Nenhuma pendencia critica encontrada no periodo.</p>}
                        </div>
                      </section>
                      <section className="instrument-management-card">
                        <p className="eyebrow">Obras / Projetos</p>
                        <h4>Resumo por obra</h4>
                        <div className="instrument-work-list">
                          {instrumentWorks.slice(0, 6).map((work) => (
                            <div key={work.key} className="instrument-work-row">
                              <strong>{work.name}</strong>
                              <span>{work.issues} issues • {work.team} colaboradores • {work.materialCount} materiais</span>
                              <small>{formatDate(work.start)} a {formatDate(work.end)}</small>
                            </div>
                          ))}
                          {!instrumentWorks.length && <p className="muted-text">Nenhuma obra identificada nos issues de instrumentacao.</p>}
                        </div>
                      </section>
                    </div>
                  </>
                )}

                {activeInstrumentSection === 'agenda' && (
                  <>
                <div className="instrumentos-filters">
                  <label>
                    Visualizacao
                    <select value={instrumentViewMode} onChange={(event) => setInstrumentViewMode(event.target.value)}>
                      <option value="week">Semana</option>
                      <option value="month">Mes</option>
                      <option value="quarter">3 meses</option>
                      <option value="custom">Periodo personalizado</option>
                    </select>
                  </label>
                  <label>
                    Data inicial
                    <input type="date" value={instrumentStartDate} onChange={(event) => setInstrumentStartDate(event.target.value)} />
                  </label>
                  <label>
                    Data final
                    <input type="date" value={instrumentEndDate} onChange={(event) => setInstrumentEndDate(event.target.value)} />
                  </label>
                  <label>
                    Obra / Empresa
                    <select value={instrumentWorkFilter} onChange={(event) => setInstrumentWorkFilter(event.target.value)}>
                      <option value="all">Todas</option>
                      {instrumentFilterOptions.works.map((item) => <option key={item} value={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Colaborador
                    <select value={instrumentResourceFilter} onChange={(event) => setInstrumentResourceFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {instrumentFilterOptions.resources.map((item) => <option key={item.value} value={item.value}>{item.label}</option>)}
                    </select>
                  </label>
                  <label>
                    Funcao operacional
                    <select value={instrumentFunctionFilter} onChange={(event) => setInstrumentFunctionFilter(event.target.value)}>
                      <option value="all">Todas</option>
                      {instrumentFilterOptions.functions.map((item) => <option key={item} value={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Status
                    <select value={instrumentStatusFilter} onChange={(event) => setInstrumentStatusFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {instrumentFilterOptions.statuses.map((item) => <option key={item} value={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Localizacao
                    <select value={instrumentLocationFilter} onChange={(event) => setInstrumentLocationFilter(event.target.value)}>
                      <option value="all">Todas</option>
                      {instrumentFilterOptions.locations.map((item) => <option key={item} value={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Tipo de recurso
                    <select value={instrumentResourceTypeFilter} onChange={(event) => setInstrumentResourceTypeFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {instrumentFilterOptions.resourceTypes.map((item) => <option key={item} value={item}>{item}</option>)}
                    </select>
                  </label>
                  <label>
                    Material
                    <select value={instrumentMaterialFilter} onChange={(event) => setInstrumentMaterialFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {instrumentFilterOptions.materials.map((item) => <option key={item} value={item}>{item}</option>)}
                    </select>
                  </label>
                  <button
                    type="button"
                    className="ghost-button"
                    onClick={() => {
                      setInstrumentWorkFilter('all');
                      setInstrumentResourceFilter('all');
                      setInstrumentFunctionFilter('all');
                      setInstrumentStatusFilter('all');
                      setInstrumentLocationFilter('all');
                      setInstrumentResourceTypeFilter('all');
                      setInstrumentMaterialFilter('all');
                    }}
                  >
                    Limpar filtros
                  </button>
                </div>

                <div className="instrumentos-indicators">
                  <div><span>Colaboradores ativos</span><strong>{instrumentationIndicators.activeResources}</strong></div>
                  <div><span>Obras/atividades</span><strong>{instrumentationIndicators.works}</strong></div>
                  <div><span>Alocacoes</span><strong>{instrumentationIndicators.allocations}</strong></div>
                  <div><span>Materiais</span><strong>{instrumentationIndicators.materials}</strong></div>
                  <div className={instrumentationIndicators.conflicts ? 'danger' : ''}><span>Conflitos</span><strong>{instrumentationIndicators.conflicts}</strong></div>
                  <div><span>Sem inicio</span><strong>{instrumentationIndicators.missingStart}</strong></div>
                  <div><span>Sem vencimento</span><strong>{instrumentationIndicators.missingEnd}</strong></div>
                  <div><span>Sem responsavel</span><strong>{instrumentationIndicators.missingResponsible}</strong></div>
                  <div className={instrumentationIndicators.overdue ? 'danger' : ''}><span>Em atraso</span><strong>{instrumentationIndicators.overdue}</strong></div>
                </div>

                <div className="instrument-calendar-board">
                  {instrumentMonthBoards.map((month) => (
                    <section key={month.key} className="instrument-month-card">
                      <header>
                        <strong>{month.label}</strong>
                        <span>{month.days.length} dias</span>
                      </header>
                      <div className="instrument-month-weekdays">
                        {['D', 'S', 'T', 'Q', 'Q', 'S', 'S'].map((weekday, index) => <span key={`${month.key}:${weekday}:${index}`}>{weekday}</span>)}
                      </div>
                      <div className="instrument-month-grid">
                        {Array.from({ length: month.leadingBlanks }).map((_, index) => <span key={`${month.key}:blank:${index}`} className="calendar-empty" />)}
                        {month.days.map((day) => {
                          const dayKey = formatInputDate(day);
                          const dayAllocations = month.allocations[dayKey] || [];
                          const isSelected = selectedInstrumentAllocation && dayAllocations.some((allocation) => allocation.id === selectedInstrumentAllocation.id);
                          return (
                            <button
                              key={dayKey}
                              type="button"
                              className={`instrument-calendar-day ${dayAllocations.length ? 'has-allocation' : ''} ${dayAllocations.length > 1 ? 'has-conflict' : ''} ${isSelected ? 'selected' : ''}`}
                              onClick={() => dayAllocations[0] && setSelectedInstrumentAllocationId(dayAllocations[0].id)}
                              disabled={!dayAllocations.length}
                              title={dayAllocations.map((allocation) => `${allocation.title} - ${allocation.material || 'Sem material'}`).join('\n')}
                            >
                              <span>{day.getDate()}</span>
                              {dayAllocations.length > 0 && <small>{dayAllocations.length}</small>}
                            </button>
                          );
                        })}
                      </div>
                    </section>
                  ))}
                </div>

                <div className="instrumentos-grid-area">
                  <div className="instrumentos-matrix-wrap">
                    <table className="instrumentos-matrix">
                      <thead>
                        <tr>
                          <th>Colaborador</th>
                          {instrumentDays.map((day) => (
                            <th key={formatInputDate(day)}>
                              <span>{day.toLocaleDateString('pt-BR', { day: '2-digit', month: '2-digit' })}</span>
                              <small>{day.toLocaleDateString('pt-BR', { weekday: 'short' })}</small>
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {instrumentResources.map((resource) => (
                          <tr key={resource.key}>
                            <th>
                              <strong>{resource.name}</strong>
                              <small>{resource.operationalRole || 'Sem funcao definida'} {resource.resourceType ? `- ${resource.resourceType}` : ''}</small>
                            </th>
                            {instrumentDays.map((day) => {
                              const allocations = instrumentAllocationByDay.get(`${resource.key}:${formatInputDate(day)}`) || [];
                              return (
                                <td key={`${resource.key}:${formatInputDate(day)}`} className={allocations.length > 1 ? 'has-conflict' : allocations.length ? 'allocated' : ''}>
                                  {allocations.slice(0, 2).map((allocation) => (
                                    <button
                                      key={allocation.id}
                                      type="button"
                                      className={`instrument-allocation ${allocation.role === 'Apoio' ? 'support' : ''}`}
                                      onClick={() => setSelectedInstrumentAllocationId(allocation.id)}
                                    >
                                      <span>{allocation.companyWork}</span>
                                      {allocation.material && <small>{allocation.material}</small>}
                                    </button>
                                  ))}
                                  {allocations.length > 2 && <span className="allocation-more">+{allocations.length - 2}</span>}
                                </td>
                              );
                            })}
                          </tr>
                        ))}
                        {!instrumentResources.length && (
                          <tr>
                            <td colSpan={instrumentDays.length + 1}>Nenhum recurso encontrado nos issues de Controle de Obras - Instrumentacao.</td>
                          </tr>
                        )}
                      </tbody>
                    </table>
                  </div>

                  <aside className="instrument-detail-panel">
                    {selectedInstrumentAllocation ? (
                      <>
                        <p className="eyebrow">Issue selecionada</p>
                        <h3>{selectedInstrumentAllocation.title}</h3>
                        <dl>
                          <div><dt>Status</dt><dd>{selectedInstrumentAllocation.status}</dd></div>
                          <div><dt>Obra</dt><dd>{selectedInstrumentAllocation.companyWork}</dd></div>
                          <div><dt>Periodo</dt><dd>{formatDate(selectedInstrumentAllocation.startDate)} a {formatDate(selectedInstrumentAllocation.endDate)}</dd></div>
                          <div><dt>Responsavel</dt><dd>{selectedInstrumentAllocation.resourceName} ({selectedInstrumentAllocation.role})</dd></div>
                          <div><dt>Servico</dt><dd>{selectedInstrumentAllocation.service || 'Nao informado'}</dd></div>
                          <div><dt>Localizacao</dt><dd>{selectedInstrumentAllocation.location || 'Nao informada'}</dd></div>
                          <div><dt>Material</dt><dd>{selectedInstrumentAllocation.material || 'Nao informado'}</dd></div>
                          <div><dt>Estado do material</dt><dd>{selectedInstrumentAllocation.materialState || 'Nao informado'}</dd></div>
                          <div><dt>Causa raiz</dt><dd>{selectedInstrumentAllocation.rootCause || 'Nao informada'}</dd></div>
                          <div><dt>Posicionamento</dt><dd>{selectedInstrumentAllocation.positioning || 'Nao informado'}</dd></div>
                        </dl>
                      </>
                    ) : (
                      <div className="empty-state">
                        <strong>Selecione uma celula da matriz.</strong>
                        <span>Os detalhes do issue aparecem aqui.</span>
                      </div>
                    )}
                  </aside>
                </div>

                <div className="instrumentos-bottom-panels">
                  <section className="instrument-resource-panel">
                    <p className="eyebrow">Cadastro de recursos</p>
                    <div className="resource-registry-list">
                      {instrumentResources.slice(0, 12).map((resource) => (
                        <div key={resource.key} className="resource-registry-row">
                          <strong>{resource.name}</strong>
                          <input
                            type="text"
                            value={resource.operationalRole || ''}
                            placeholder="Funcao operacional"
                            onChange={(event) => updateInstrumentResource(resource.key, 'operationalRole', event.target.value)}
                          />
                          <input
                            type="text"
                            value={resource.base || ''}
                            placeholder="Base"
                            onChange={(event) => updateInstrumentResource(resource.key, 'base', event.target.value)}
                          />
                          <select value={resource.resourceType || 'Equipe'} onChange={(event) => updateInstrumentResource(resource.key, 'resourceType', event.target.value)}>
                            <option value="Campo">Campo</option>
                            <option value="Apoio">Apoio</option>
                            <option value="Escritorio">Escritorio</option>
                            <option value="Indisponivel">Indisponivel</option>
                          </select>
                        </div>
                      ))}
                    </div>
                  </section>
                  <section className="configuration-pending-panel">
                    <p className="eyebrow">Pendencias de configuracao</p>
                    <ul>
                      {instrumentationIndicators.missingResponsible > 0 && <li>{instrumentationIndicators.missingResponsible} issues sem responsavel principal.</li>}
                      {instrumentationIndicators.missingStart > 0 && <li>{instrumentationIndicators.missingStart} issues sem data de inicio.</li>}
                      {instrumentationIndicators.missingEnd > 0 && <li>{instrumentationIndicators.missingEnd} issues sem data de vencimento.</li>}
                      {instrumentationIndicators.conflicts > 0 && <li>{instrumentationIndicators.conflicts} dias com conflito de alocacao.</li>}
                      {!instrumentationIndicators.missingResponsible && !instrumentationIndicators.missingStart && !instrumentationIndicators.missingEnd && !instrumentationIndicators.conflicts && <li>Nenhuma pendencia critica encontrada no periodo.</li>}
                    </ul>
                  </section>
                </div>
                  </>
                )}

                {activeInstrumentSection === 'equipamentos' && (
                  <div className="instrument-management-card">
                    <p className="eyebrow">Controle de Equipamentos</p>
                    <h4>Inventario e movimentacao</h4>
                    <div className="instrument-equipment-table">
                      <table>
                        <thead>
                          <tr><th>Equipamento</th><th>Status</th><th>Obra atual</th><th>Responsavel</th><th>Retorno</th><th>Calibracao</th></tr>
                        </thead>
                        <tbody>
                          {instrumentEquipmentItems.map((item) => (
                            <tr key={`${item.issue.id}:${item.code}`} className={item.returnOverdue || item.calibrationOverdue ? 'attention-row' : ''}>
                              <td><strong>{item.name}</strong><small>{item.code}</small></td>
                              <td>{item.status}</td>
                              <td>{item.obra}</td>
                              <td>{item.responsible}</td>
                              <td>{formatDate(item.returnDate)}</td>
                              <td>{formatDate(item.calibrationDate)}</td>
                            </tr>
                          ))}
                          {!instrumentEquipmentItems.length && (
                            <tr><td colSpan="6">Nenhum equipamento identificado nos campos personalizados dos issues. A tela ja esta preparada para cadastro/movimentacao.</td></tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {activeInstrumentSection === 'materiais' && (
                  <div className="instrument-management-card">
                    <p className="eyebrow">Materiais e Compras</p>
                    <h4>Lista de materiais da obra e saneamento ABNT NBR 15965</h4>
                    <div className="instrumentos-indicators compact">
                      <div><span>Total de itens</span><strong>{instrumentMaterialItems.length}</strong></div>
                      <div><span>Pendentes de compra</span><strong>{instrumentPurchaseStats.pending}</strong></div>
                      <div><span>Em cotacao</span><strong>{instrumentPurchaseStats.quoted}</strong></div>
                      <div><span>Recebidos</span><strong>{instrumentPurchaseStats.received}</strong></div>
                    </div>
                    <div className="instrument-equipment-table">
                      <table>
                        <thead>
                          <tr><th>Obra</th><th>Codigo</th><th>Material</th><th>Unidade</th><th>Quantidade</th><th>Status de compra</th></tr>
                        </thead>
                        <tbody>
                          {instrumentMaterialItems.map((item) => (
                            <tr key={`${item.issue.id}:${item.code}:${item.material}`} className={!item.classified ? 'pending-row' : ''}>
                              <td>{item.obra}</td>
                              <td>{item.code}</td>
                              <td><strong>{item.material}</strong><small>{item.classified ? 'Classificado' : 'Sem classificacao padronizada'}</small></td>
                              <td>{item.unit}</td>
                              <td>{item.quantity}</td>
                              <td>{item.compraStatus}</td>
                            </tr>
                          ))}
                          {!instrumentMaterialItems.length && (
                            <tr><td colSpan="6">Nenhum material identificado nos issues. O modulo esta preparado para importar listas de materiais por obra.</td></tr>
                          )}
                        </tbody>
                      </table>
                    </div>
                  </div>
                )}

                {activeInstrumentSection === 'obras' && (
                  <div className="instrument-work-cards">
                    {instrumentWorks.map((work) => (
                      <article key={work.key} className="instrument-work-card">
                        <p className="eyebrow">Ficha 360 da obra</p>
                        <h4>{work.name}</h4>
                        <dl>
                          <div><dt>Issues vinculados</dt><dd>{work.issues}</dd></div>
                          <div><dt>Equipe alocada</dt><dd>{work.team}</dd></div>
                          <div><dt>Apoios</dt><dd>{work.supportCount}</dd></div>
                          <div><dt>Materiais</dt><dd>{work.materialCount}</dd></div>
                          <div><dt>Periodo</dt><dd>{formatDate(work.start)} a {formatDate(work.end)}</dd></div>
                        </dl>
                      </article>
                    ))}
                    {!instrumentWorks.length && <div className="empty-state"><strong>Nenhuma obra identificada.</strong><span>Preencha Empresa+Obra nos issues de instrumentacao para montar esta visao.</span></div>}
                  </div>
                )}

                {activeInstrumentSection === 'cadastros' && (
                  <div className="instrumentos-bottom-panels">
                    <section className="instrument-resource-panel">
                      <p className="eyebrow">Recursos humanos</p>
                      <h4>Cadastro local de colaboradores</h4>
                      <div className="resource-registry-list">
                        {instrumentResources.slice(0, 16).map((resource) => (
                          <div key={resource.key} className="resource-registry-row">
                            <strong>{resource.name}</strong>
                            <input type="text" value={resource.operationalRole || ''} placeholder="Funcao operacional" onChange={(event) => updateInstrumentResource(resource.key, 'operationalRole', event.target.value)} />
                            <input type="text" value={resource.base || ''} placeholder="Base" onChange={(event) => updateInstrumentResource(resource.key, 'base', event.target.value)} />
                            <select value={resource.resourceType || 'Equipe'} onChange={(event) => updateInstrumentResource(resource.key, 'resourceType', event.target.value)}>
                              <option value="Campo">Campo</option>
                              <option value="Apoio">Apoio</option>
                              <option value="Escritorio">Escritorio</option>
                              <option value="Indisponivel">Indisponivel</option>
                            </select>
                          </div>
                        ))}
                      </div>
                    </section>
                    <section className="instrument-resource-panel">
                      <p className="eyebrow">Catalogo padronizado</p>
                      <h4>Materiais ABNT NBR 15965</h4>
                      <p className="muted-text">Estrutura preparada para vincular descricoes das obras ao catalogo interno G5, codigo ABNT NBR 15965, unidade padrao, categoria e especificacao minima.</p>
                    </section>
                  </div>
                )}

                {activeInstrumentSection === 'pendencias' && (
                  <div className="instrument-management-card">
                    <p className="eyebrow">Pendencias e Alertas</p>
                    <h4>Painel centralizado de criticidade</h4>
                    <div className="instrument-alert-list wide">
                      {instrumentAlerts.map((alert, index) => (
                        <div key={`${alert.area}:${alert.level}:${index}`} className={`instrument-alert-item level-${normalizeText(alert.level)}`}>
                          <span>{alert.level}</span>
                          <strong>{alert.area}</strong>
                          <p>{alert.text}</p>
                        </div>
                      ))}
                      {!instrumentAlerts.length && <p className="muted-text">Nenhuma pendencia critica encontrada.</p>}
                    </div>
                  </div>
                )}

                {activeInstrumentSection === 'relatorios' && (
                  <div className="instrument-management-card">
                    <p className="eyebrow">Relatorios / Exportacoes</p>
                    <h4>Base preparada para exportacoes</h4>
                    <div className="instrument-report-grid">
                      {['Matriz de alocacao de equipe', 'Equipamentos por obra', 'Equipamentos vencidos / calibracao', 'Materiais por obra', 'Compras pendentes', 'Pendencias criticas', 'Consolidado da obra', 'Disponibilidade de equipe'].map((report) => (
                        <button key={report} type="button" className="ghost-button" disabled>{report}</button>
                      ))}
                    </div>
                  </div>
                )}
              </section>
            </div>
          </section>
        )}

        {user && selectedProjectId && selectedBusinessUnit !== 'G5 Instrumentos' && (
          <>
            {!activeModule && (
              <section className="module-panel">
                <div className="module-home-top">
                  <div className="module-home-intro">
                    <p className="module-user-line">
                      <span className="module-user-badge" aria-hidden="true">👤</span>
                      Usuário autenticado: <strong>{user.name || user.email || 'Usuário Autodesk'}</strong>
                    </p>
                    <div className="module-heading module-heading-compact">
                      <span aria-hidden="true" />
                      <div>
                        <h2>Escolha uma área de trabalho</h2>
                        <p>Acesse os módulos disponíveis para acompanhamento, comunicação e gestão do projeto.</p>
                      </div>
                    </div>
                  </div>

                  <div className="module-context-card module-context-card-compact">
                    <div className="module-context-field">
                      <span>Hub</span>
                      <strong>{selectedHub?.name || 'Hub selecionado'}</strong>
                    </div>
                    <div className="module-context-field">
                      <span>Projeto</span>
                      <strong>{selectedProject?.name || selectedProjectId}</strong>
                    </div>
                    <button type="button" className="module-project-switch" onClick={() => setProjectPickerOpen(true)}>
                      Trocar projeto
                    </button>
                  </div>
                </div>
                <div className={`module-home-layout ${centralBimMenuOpen ? 'menu-open' : 'menu-collapsed'}`}>
                  <aside className="central-bim-menu" aria-label="Menu lateral Central BIM">
                    <header>
                      <button
                        type="button"
                        className="central-bim-menu-toggle"
                        onClick={() => setCentralBimMenuOpen((isOpen) => !isOpen)}
                        aria-label="Abrir menu Central BIM"
                        aria-expanded={centralBimMenuOpen}
                      >
                        <span aria-hidden="true" />
                        <span aria-hidden="true" />
                        <span aria-hidden="true" />
                      </button>
                      {centralBimMenuOpen && <strong>Central BIM</strong>}
                    </header>
                    {centralBimMenuOpen && (
                      <div className="central-bim-menu-items">
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('im'); resetImModule(); setImMode(''); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="modules" /></span>
                          <span className="module-copy">
                            <strong>Gestão de Issues</strong>
                            <span>Crie e atualize issues no ACC com base na planilha ACS Build.</span>
                          </span>
                        </button>
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('interface'); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="request" /></span>
                          <span className="module-copy">
                            <strong>Abrir Solicitação</strong>
                            <span>Registre demandas de informação no ACC com categoria, tipo e campos personalizados.</span>
                          </span>
                        </button>
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('project-management'); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="project" /></span>
                          <span className="module-copy">
                            <strong>Gestão do projeto</strong>
                            <span>Marcos contratuais e valores usando somente issues visíveis para o usuário logado.</span>
                          </span>
                        </button>
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('eap-import'); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="eap" /></span>
                          <span className="module-copy">
                            <strong>Importar EAP</strong>
                            <span>Importe a planilha ACS Build, valide os dados e prepare a criação de issues.</span>
                          </span>
                        </button>
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('qualidade'); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="overview" /></span>
                          <span className="module-copy">
                            <strong>Qualidade BIM</strong>
                            <span>Painel de Qualidade dos Issues com auditoria, IQI e governança de dados ACC.</span>
                          </span>
                        </button>
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('bim-awp'); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="bim" /></span>
                          <span className="module-copy">
                            <strong>Coordenação BIM</strong>
                            <span>Modelo BIM, workfaces, validação, qualidade e publicação.</span>
                          </span>
                        </button>
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('meetings'); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="meeting" /></span>
                          <span className="module-copy">
                            <strong>Reunioes</strong>
                            <span>Modelos para Design Review, reuniao tecnica e reuniao gerencial.</span>
                          </span>
                        </button>
                        <button className="central-bim-menu-item" type="button" onClick={() => { setActiveModule('interfaces-pendencias'); setCentralBimMenuOpen(false); }}>
                          <span className="module-icon" aria-hidden="true"><UiIcon name="interfaces" /></span>
                          <span className="module-copy">
                            <strong>Interfaces e Pendências</strong>
                            <span>Pendências por disciplina, impactos e encaminhamentos dos issues.</span>
                          </span>
                        </button>
                      </div>
                    )}
                  </aside>

                <div className="module-grid">
                  <button className="module-button module-button-primary" type="button" onClick={() => setActiveModule('coordination')}>
                    <span className="module-icon" aria-hidden="true"><UiIcon name="timeline" /></span>
                    <span className="module-copy">
                      <strong>Linha do Tempo - Issues</strong>
                      <span>Acompanhe issues do projeto em uma linha do tempo mensal.</span>
                    </span>
                    <span className="module-arrow" aria-hidden="true">›</span>
                  </button>
                  <button className="module-button module-button-links" type="button" onClick={() => setActiveModule('links')}>
                    <span className="module-icon" aria-hidden="true"><UiIcon name="overview" /></span>
                    <span className="module-copy">
                      <strong>Mapa das atividades</strong>
                      <span>Visualize relações, referências e dependências entre issues do projeto.</span>
                    </span>
                    <span className="module-arrow" aria-hidden="true">›</span>
                  </button>
                  <button className="module-button module-button-documents" type="button" onClick={openDocumentsModule}>
                    <span className="module-icon" aria-hidden="true"><UiIcon name="documents" /></span>
                    <span className="module-copy">
                      <strong>Documentos emitidos</strong>
                      <span>Compare a lista mestre com os arquivos publicados no CDE.</span>
                    </span>
                    <span className="module-arrow" aria-hidden="true">›</span>
                  </button>
                  <button className="module-button module-button-impact" type="button" onClick={() => setActiveModule('impact')}>
                    <span className="module-icon" aria-hidden="true"><UiIcon name="impact" /></span>
                    <span className="module-copy">
                      <strong>Impactos no projeto</strong>
                      <span>Analise impactos em prazo, escopo e medição por criticidade e vencimento.</span>
                    </span>
                    <span className="module-arrow" aria-hidden="true">›</span>
                  </button>
                  <button className="module-button module-button-primary" type="button" onClick={() => setActiveModule('cronograma-im')}>
                    <span className="module-icon" aria-hidden="true"><UiIcon name="timeline" /></span>
                    <span className="module-copy">
                      <strong>EAP</strong>
                      <span>Cronograma visual analítico baseado em Start Date, Due Date e referências ACC.</span>
                    </span>
                    <span className="module-arrow" aria-hidden="true">›</span>
                  </button>
                  <button className="module-button module-button-schedule" type="button" onClick={() => setActiveModule('schedule')}>
                    <span className="module-icon" aria-hidden="true"><UiIcon name="timeline" /></span>
                    <span className="module-copy">
                      <strong>Visão geral do projeto</strong>
                      <span>Motor de marcos, tramitação, riscos, dependências e avanço executivo por Issues.</span>
                    </span>
                    <span className="module-arrow" aria-hidden="true">›</span>
                  </button>
                </div>

                <section className="home-planner" aria-label="Planner de Entregas">
                  <div className="home-planner-header">
                    <div>
                      <p className="eyebrow">PLANEJAMENTO</p>
                      <h3>Planner de Entregas</h3>
                      <p>Visualização mensal dos marcos, entregas e desenvolvimentos previstos para o projeto.</p>
                    </div>
                    <div className="home-planner-actions">
                      <div className="planner-legend" aria-label="Legenda do planner">
                        <span><i className="planner-dot planner-dot-delivery" /> Marco Contratual / Entrega</span>
                        <span><i className="planner-dot planner-dot-development" /> Desenvolvimento</span>
                      </div>
                      <div className="planner-nav" aria-label="Navegação do planner">
                        <button type="button" onClick={() => setPlannerStartMonth((current) => addMonths(current, -3))} aria-label="Voltar 3 meses">‹</button>
                        <button type="button" onClick={() => setPlannerStartMonth(getMonthStart(new Date()))}>Hoje</button>
                        <button type="button" onClick={() => setPlannerStartMonth((current) => addMonths(current, 3))} aria-label="Avançar 3 meses">›</button>
                      </div>
                    </div>
                  </div>

                  <div className="planner-summary-strip">
                    <span><strong>{plannerEventTotals.deliveries}</strong> entregas/marcos mapeados</span>
                    <span><strong>{plannerEventTotals.developments}</strong> desenvolvimentos mapeados</span>
                    {issuesLoading && <span>Atualizando issues do projeto...</span>}
                  </div>

                  <div className="planner-months-grid">
                    {plannerVisibleMonths.map((monthDate) => (
                      <article className="planner-month-card" key={formatInputDate(monthDate)}>
                        <h4>{formatPlannerMonthTitle(monthDate)}</h4>
                        <div className="planner-weekdays" aria-hidden="true">
                          {['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'].map((weekday) => <span key={weekday}>{weekday}</span>)}
                        </div>
                        <div className="planner-days-grid">
                          {buildPlannerMonthDays(monthDate).map((day) => {
                            const dateKey = formatInputDate(day);
                            const events = plannerEventsByDate.get(dateKey) || [];
                            const hasDelivery = events.some((event) => event.kind === 'delivery');
                            const hasDevelopment = events.some((event) => event.kind === 'development');
                            const isCurrentMonth = day.getMonth() === monthDate.getMonth();
                            const isToday = dateKey === formatInputDate(new Date());
                            const isSelected = selectedPlannerDateKey === dateKey;
                            return (
                              <button
                                key={dateKey}
                                type="button"
                                className={`planner-day ${isCurrentMonth ? '' : 'planner-day-muted'} ${isToday ? 'planner-day-today' : ''} ${isSelected ? 'planner-day-selected' : ''}`}
                                onClick={() => setSelectedPlannerDateKey(events.length ? dateKey : '')}
                                disabled={!events.length}
                                title={events.length ? `${events.length} item(ns) em ${formatDate(day)}` : ''}
                              >
                                <span className="planner-day-number">{day.getDate()}</span>
                                <span className="planner-day-dots" aria-hidden="true">
                                  {hasDelivery && <i className="planner-dot planner-dot-delivery" />}
                                  {hasDevelopment && <i className="planner-dot planner-dot-development" />}
                                </span>
                              </button>
                            );
                          })}
                        </div>
                      </article>
                    ))}
                  </div>

                  <div className="planner-day-detail">
                    {selectedPlannerDateKey && selectedPlannerEvents.length ? (
                      <>
                        <div className="planner-day-detail-title">
                          <strong>{formatDate(selectedPlannerDateKey)}</strong>
                          <span>{selectedPlannerEvents.length} item(ns)</span>
                        </div>
                        <div className="planner-event-list">
                          {selectedPlannerEvents.map((event) => (
                            <button
                              key={`${event.issue.id || event.issue.displayId}-${event.kind}`}
                              type="button"
                              className="planner-event-item"
                              onClick={() => openPlannerIssueInKanban(event.issue.id)}
                            >
                              <i className={`planner-dot ${event.kind === 'delivery' ? 'planner-dot-delivery' : 'planner-dot-development'}`} />
                              <span>
                                <strong>{event.title}</strong>
                                <small>{event.label} · {event.category} · {event.status}</small>
                              </span>
                            </button>
                          ))}
                        </div>
                      </>
                    ) : (
                      <p className="planner-empty-message">
                        Selecione um dia sinalizado para visualizar os issues previstos. Issues sem data válida não são exibidos no planner.
                      </p>
                    )}
                  </div>
                </section>
                </div>
              </section>
            )}

            {activeModule && (
          <section className="monitor-panel">
            {activeModule !== 'cronograma-im' && activeModule !== 'schedule' && (
            <div className="section-heading">
              <div>
                <p className="eyebrow">
                  {activeModule === 'interface'
                    ? 'ABERTURA'
                    : activeModule === 'links'
                      ? 'VÍNCULOS'
                    : activeModule === 'interfaces-pendencias'
                      ? 'INTERFACES'
                    : activeModule === 'documents'
                        ? 'DOCUMENTOS'
                        : activeModule === 'eap-import'
                          ? 'IMPORTACAO'
                        : activeModule === 'im'
                          ? 'IM'
                        : activeModule === 'impact'
                          ? 'ANÁLISE'
                        : activeModule === 'bim-awp'
                          ? 'BIM + AWP'
                        : activeModule === 'project-management'
                          ? 'PROJETO'
                        : activeModule === 'meetings'
                          ? 'REUNIOES'
                        : activeModule === 'qualidade'
                          ? 'QUALIDADE'
                        : 'GESTÃO'}
                </p>
                <h2>
                  {activeModule === 'interface'
                    ? 'SOLICITAR INFORMAÇÃO'
                    : activeModule === 'links'
                      ? 'KANBAN DE MARCOS'
                    : activeModule === 'interfaces-pendencias'
                      ? 'INTERFACES E PENDÊNCIAS'
                    : activeModule === 'documents'
                        ? 'DOCUMENTOS EMITIDOS'
                        : activeModule === 'eap-import'
                          ? 'IMPORTAR ISSUES ACS BUILD'
                        : activeModule === 'im'
                          ? 'GESTÃO DE ISSUES'
                        : activeModule === 'impact'
                          ? 'IMPACTOS NO PROJETO'
                        : activeModule === 'bim-awp'
                          ? 'COORDENAÇÃO BIM'
                        : activeModule === 'project-management'
                          ? 'GESTÃO DO PROJETO'
                        : activeModule === 'meetings'
                          ? 'MODELOS DE REUNIAO'
                        : activeModule === 'qualidade'
                          ? 'PAINEL DE QUALIDADE DOS ISSUES'
                        : 'LINHA DO TEMPO - ISSUES'}
                </h2>
                <p className="selected-project-name">
                  Projeto selecionado: <strong>{selectedProject?.name || selectedProjectId}</strong>
                </p>
              </div>
              <div className="section-actions">
                <button className="ghost-button" type="button" onClick={() => setActiveModule('')}>
                  Módulos
                </button>
                <button
                  className="ghost-button"
                  type="button"
                  onClick={refreshCurrentModule}
                >
                  Atualizar
                </button>
              </div>
            </div>
            )}

            {activeModule === 'interface' && (
              <form className="create-issue-panel" onSubmit={createInterfaceIssue}>
                <div>
                  <p className="eyebrow">Dados principais</p>
                  <h3>INFORMAÇÃO TÉCNICA</h3>
                </div>
                <label>
                  Categoria *
                  <select
                    required
                    value={newInterfaceIssue.issueTypeId}
                    onChange={(event) =>
                      setNewInterfaceIssue((issue) => ({
                        ...issue,
                        issueTypeId: event.target.value,
                        issueSubtypeId: '',
                        customAttributes: {}
                      }))
                    }
                  >
                    <option value="">Selecione a categoria</option>
                    {issueCategories.map((category) => (
                      <option key={category.id} value={category.id}>
                        {category.title}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  Tipo
                  <select
                    value={newInterfaceIssue.issueSubtypeId}
                    onChange={(event) =>
                      setNewInterfaceIssue((issue) => ({
                        ...issue,
                        issueSubtypeId: event.target.value,
                        customAttributes: {}
                      }))
                    }
                  >
                    <option value="">Selecione o tipo</option>
                    {issueSubtypes.map((subtype) => (
                      <option key={subtype.id} value={subtype.id}>
                        {subtype.title}
                      </option>
                    ))}
                  </select>
                </label>
                <label>
                  Título *
                  <input
                    required
                    value={newInterfaceIssue.title}
                    onChange={(event) => setNewInterfaceIssue((issue) => ({ ...issue, title: event.target.value }))}
                    placeholder="Ex.: Compatibilizar interferência entre disciplinas"
                  />
                </label>
                <label>
                  Prazo
                  <input
                    type="date"
                    value={newInterfaceIssue.dueDate}
                    onChange={(event) => setNewInterfaceIssue((issue) => ({ ...issue, dueDate: event.target.value }))}
                  />
                </label>
                <label className="wide-field">
                  Atribuir a
                  <select
                    value={newInterfaceIssue.assignedTo}
                    onChange={(event) => setNewInterfaceIssue((issue) => ({ ...issue, assignedTo: event.target.value }))}
                  >
                    <option value="">Sem responsável</option>
                    {projectUsers.map((projectUser) => (
                      <option key={projectUser.id} value={projectUser.id}>
                        {projectUser.name}
                        {projectUser.email ? ` - ${projectUser.email}` : ''}
                      </option>
                    ))}
                  </select>
                </label>
                <label className="wide-field">
                  Descrição
                  <textarea
                    value={newInterfaceIssue.description}
                    onChange={(event) => setNewInterfaceIssue((issue) => ({ ...issue, description: event.target.value }))}
                    placeholder="Descreva a interface, disciplinas envolvidas, pendência e ação esperada."
                  />
                </label>
                {newInterfaceIssue.issueTypeId && applicableCustomFieldDefinitions.length > 0 && (
                  <div className="new-custom-fields wide-field">
                    <p className="label">Informações:</p>
                    <div className="new-custom-grid">
                      {applicableCustomFieldDefinitions.map((field) => (
                        <label key={field.id}>
                          {field.name}{field.required ? ' *' : ''}
                          {field.options?.length > 0 ? (
                            <select
                              required={field.required}
                              value={newInterfaceIssue.customAttributes[field.id] || ''}
                              onChange={(event) => updateNewCustomAttribute(field.id, event.target.value)}
                            >
                              <option value="">{field.required ? 'Selecione uma opção' : 'Não informado'}</option>
                              {field.options.map((option) => (
                                <option key={option.id} value={option.id}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                          ) : (
                            <input
                              required={field.required}
                              value={newInterfaceIssue.customAttributes[field.id] || ''}
                              onChange={(event) => updateNewCustomAttribute(field.id, event.target.value)}
                              placeholder={field.required ? 'Campo obrigatório' : 'Preencher se necessário'}
                            />
                          )}
                        </label>
                      ))}
                    </div>
                  </div>
                )}
                <button className="primary-button" type="submit">
                  Criar no ACC
                </button>
              </form>
            )}

            {activeModule === 'links' && (
              <section className="link-flow-panel">
                <div className="link-flow-intro">
                  <div>
                    <p className="eyebrow">Quadro operacional</p>
                    <h3>KANBAN POR MARCO CONTRATUAL</h3>
                  </div>
                  <span>{issues.length} issues no projeto</span>
                </div>

                <div className="kanban-quick-actions" aria-label="Ações rápidas do Kanban">
                  <form className="kanban-create-panel kanban-column-panel kanban-compact-panel" onSubmit={createKanbanMarcoColumn}>
                    <div className="kanban-create-heading">
                      <div>
                        <p className="eyebrow">Nova coluna</p>
                        <h4>Criar Marco Contratual</h4>
                        <span>Cria a coluna do quadro e usa o valor para gravar o campo Marco Contratual ao mover cards.</span>
                      </div>
                      <button className="primary-button" type="submit">
                        Criar coluna
                      </button>
                    </div>
                    <div className="kanban-create-grid kanban-column-create-grid">
                      <label className="wide-field">
                        Marco Contratual *
                        <input value={kanbanColumnDraft.title} onChange={(event) => updateKanbanColumnDraft('title', event.target.value)} placeholder="Ex.: Marco - Desenvolvimento do Projeto Básico" required />
                      </label>
                      <label>
                        Código do Marco
                        <input value={kanbanColumnDraft.code} onChange={(event) => updateKanbanColumnDraft('code', event.target.value)} placeholder="Ex.: M03" />
                      </label>
                    </div>
                    {kanbanColumnMessage && <p className="kanban-column-message">{kanbanColumnMessage}</p>}
                  </form>

                  <form className="kanban-create-panel kanban-card-create-panel kanban-compact-panel" onSubmit={createKanbanCard}>
                    <div className="kanban-create-heading">
                      <div>
                        <p className="eyebrow">Novo card</p>
                        <h4>Criar issue no ACC</h4>
                        <span>Registra uma nova pendência de Interface e Coordenação Multidisciplinar.</span>
                      </div>
                      <button className="primary-button" type="submit" disabled={kanbanCreating}>
                        {kanbanCreating ? 'Criando...' : 'Criar card'}
                      </button>
                    </div>
                    <div className="kanban-create-grid">
                      <label>
                        Marco Contratual *
                        <select value={kanbanDraft.marcoId} onChange={(event) => updateKanbanDraft('marcoId', event.target.value)} required>
                          <option value="">Selecione um Marco Contratual</option>
                          {linkKanbanColumns.filter((column) => column.canReceiveCards).map((column) => (
                            <option key={column.id} value={column.id}>{getKanbanColumnOptionLabel(column)}</option>
                          ))}
                        </select>
                      </label>
                      <label>
                        Tipo do issue *
                        <select value={kanbanDraft.type} onChange={(event) => updateKanbanDraft('type', event.target.value)} required>
                          {kanbanIssueTypes.map((type) => <option key={type} value={type}>{type}</option>)}
                        </select>
                      </label>
                      <label className="wide-field">
                        Título *
                        <input value={kanbanDraft.title} onChange={(event) => updateKanbanDraft('title', event.target.value)} placeholder="Ex.: Validar interferência entre disciplinas" required />
                      </label>
                      <label>
                        Responsável *
                        <select value={kanbanDraft.assignedTo} onChange={(event) => updateKanbanDraft('assignedTo', event.target.value)} required>
                          <option value="">Selecione</option>
                          {projectUsers.map((projectUser) => (
                            <option key={projectUser.id} value={projectUser.id}>
                              {projectUser.name}{projectUser.email ? ` - ${projectUser.email}` : ''}
                            </option>
                          ))}
                        </select>
                      </label>
                      <label>
                        Data prevista *
                        <input type="date" value={kanbanDraft.dueDate} onChange={(event) => updateKanbanDraft('dueDate', event.target.value)} required />
                      </label>
                      <label>
                        Disciplina *
                        <select value={kanbanDraft.disciplina} onChange={(event) => updateKanbanDraft('disciplina', event.target.value)} required>
                          <option value="">Selecione</option>
                          {kanbanDisciplineOptions.map((option) => (
                            <option key={option.id} value={option.id}>{option.label}</option>
                          ))}
                        </select>
                      </label>
                      <label className="wide-field kanban-description-field">
                        Descrição
                        <textarea value={kanbanDraft.description} onChange={(event) => updateKanbanDraft('description', event.target.value)} placeholder="Contexto, pendência, decisão esperada ou disciplinas envolvidas." />
                      </label>
                    </div>
                  </form>
                </div>

                <div className="link-flow-filters">
                  <label>
                    Categoria *
                    <select value={linkCategoryFilter} onChange={(event) => setLinkCategoryFilter(event.target.value)}>
                      <option value="all">Todas as categorias</option>
                      {linkCategoryOptions.map((category) => (
                        <option key={category} value={category}>
                          {category}
                        </option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Tipo
                    <select value={linkTypeFilter} onChange={(event) => setLinkTypeFilter(event.target.value)}>
                      <option value="all">Todos os tipos</option>
                      {linkFilterOptions.types.map((type) => <option key={type} value={type}>{type}</option>)}
                    </select>
                  </label>
                  <label>
                    Status
                    <select value={linkStatusFilter} onChange={(event) => setLinkStatusFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {statusOptions.map((status) => <option key={status.value} value={status.value}>{status.label}</option>)}
                    </select>
                  </label>
                  <label>
                    Código do Marco
                    <select value={linkMarcoFilter} onChange={(event) => setLinkMarcoFilter(event.target.value)}>
                      <option value="all">Todas</option>
                      {linkFilterOptions.codigosMarco.map((codigo) => <option key={codigo} value={codigo}>{codigo}</option>)}
                    </select>
                  </label>
                  <label>
                    Responsável
                    <select value={linkResponsibleFilter} onChange={(event) => setLinkResponsibleFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {linkFilterOptions.responsibles.map((responsible) => <option key={responsible} value={responsible}>{responsible}</option>)}
                    </select>
                  </label>
                  <label>
                    Campo personalizado
                    <select value={linkCustomFieldFilter} onChange={(event) => setLinkCustomFieldFilter(event.target.value)}>
                      <option value="all">Selecionar</option>
                      {linkFilterOptions.customFields.map((field) => <option key={field} value={field}>{field}</option>)}
                    </select>
                  </label>
                  <button className="ghost-button" type="button" onClick={() => refreshCurrentModule()}>
                    Atualizar quadro
                  </button>
                  <span className="kanban-autosave-badge">Salvamento automático no ACC</span>
                  <button className="ghost-button" type="button" onClick={exportIssueLinksManagementSheet}>
                    <UiIcon name="export" /> Exportar planilha de gerenciamento
                  </button>
                </div>

                {issues.length === 0 && (
                  <div className="empty-state">
                    <strong>Nenhuma issue foi carregada. Clique em Atualizar para buscar os issues do projeto selecionado.</strong>
                  </div>
                )}

                {issues.length > 0 && linkKanbanColumns.length === 0 && (
                  <div className="empty-state">
                    <strong>Nenhuma issue encontrada para esta categoria.</strong>
                    <span>Escolha outra combinação de filtros para visualizar os pacotes/marcos disponíveis.</span>
                  </div>
                )}

                {linkKanbanColumns.length > 0 && (
                  <>
                    {kanbanDisciplineOptions.length > 0 && (
                      <div className="kanban-discipline-legend" aria-label="Legenda de disciplinas">
                        {kanbanDisciplineOptions.map((option) => (
                          <span key={option.id} style={{ '--discipline-color': getDisciplineColor(option.label) }}>
                            <i /> {option.label}
                          </span>
                        ))}
                      </div>
                    )}

                    <div className="kanban-drag-instructions">
                      <strong>Movimentação no quadro:</strong>
                      <span>arraste o card para uma coluna de Marco Contratual. Ao soltar, o campo personalizado <b>Marco Contratual</b> será gravado automaticamente no ACC.</span>
                    </div>

                    <div className="kanban-board" aria-label="Quadro Kanban de atividades por entrega">
                      {linkKanbanColumns.map((column) => (
                        <section
                          key={column.id}
                          className={`kanban-column ${column.canReceiveCards ? '' : 'unlinked-column'} ${kanbanDragOverColumnId === column.id ? 'is-drop-target' : ''}`}
                          onDragOver={(event) => {
                            if (!column.canReceiveCards) return;
                            event.preventDefault();
                          }}
                          onDragEnter={(event) => {
                            if (!column.canReceiveCards) return;
                            event.preventDefault();
                            setKanbanDragOverColumnId(column.id);
                          }}
                          onDragLeave={(event) => {
                            if (event.currentTarget.contains(event.relatedTarget)) return;
                            setKanbanDragOverColumnId('');
                          }}
                          onDrop={(event) => handleKanbanDrop(event, column)}
                        >
                          <header className="kanban-column-header">
                            <div className="kanban-column-heading">
                              <span className="kanban-column-code">{column.code || 'Sem código'}</span>
                              <h4>{column.deliveryTitle || column.title || column.marcoContratual}</h4>
                              {getKanbanColumnSubtitle(column) && <small>{getKanbanColumnSubtitle(column)}</small>}
                              {column.issue && column.title && normalizeText(column.title) !== normalizeText(column.marcoContratual || '') && (
                                <em>{column.title}</em>
                              )}
                              {column.isManual && !column.cards.length && (
                                <button className="kanban-column-remove" type="button" onClick={() => removeKanbanManualColumn(column.id)}>Remover coluna</button>
                              )}
                            </div>
                            <div className="kanban-column-kpis">
                              <b>{column.cards.length}</b>
                              <small>{column.open} abertas • {column.overdue} atrasadas</small>
                            </div>
                          </header>

                          <div className="kanban-card-list">
                            {column.cards.length === 0 ? (
                              <p className="kanban-empty">Nenhum card vinculado a esta entrega.</p>
                            ) : (
                              column.cards.map((issue) => {
                                const discipline = getIssueDiscipline(issue);
                                return (
                                  <article
                                    key={issue.id}
                                    className={`kanban-card ${isOverdue(issue) ? 'is-overdue' : ''} ${!isOpenIssue(issue) ? 'is-closed' : ''} ${selectedKanbanIssueId === issue.id ? 'is-selected' : ''} ${kanbanDraggedIssueId === issue.id ? 'is-dragging' : ''}`}
                                    style={{ '--discipline-color': getDisciplineColor(discipline) }}
                                    draggable={Boolean(issue.id)}
                                    onClick={() => setSelectedKanbanIssueId(issue.id)}
                                    onDragStart={(event) => handleKanbanDragStart(event, issue)}
                                    onDragEnd={handleKanbanDragEnd}
                                  >
                                    <button
                                      type="button"
                                      className="kanban-card-main"
                                      onClick={(event) => {
                                        event.stopPropagation();
                                        setSelectedKanbanIssueId(issue.id);
                                      }}
                                    >
                                      <span className="kanban-type">{issue.issueType || issue.issueSubtype || 'Sem tipo'}</span>
                                      <strong>{issue.title}</strong>
                                    </button>
                                    <div className="kanban-card-meta">
                                      <span className="kanban-discipline-chip"><i /> {discipline || 'Sem disciplina'}</span>
                                      <span className="kanban-status-chip">{getIssueStatusLabel(issue.status)}</span>
                                      <span className="kanban-assignee-name">{issue.assignedTo || 'Sem responsável'}</span>
                                      <span className="kanban-date-chip">{issue.dueDate ? formatDate(issue.dueDate) : 'Sem prazo'}</span>
                                    </div>
                                    <div className="kanban-card-footer">
                                      {isOverdue(issue) && <span className="kanban-alert">Atrasada</span>}
                                      <span className="kanban-drag-hint">{kanbanMovingIssueId === issue.id ? 'Salvando no ACC...' : 'Arraste para mover'}</span>
                                    </div>
                                  </article>
                                );
                              })
                            )}
                          </div>
                        </section>
                      ))}
                    </div>
                    <section className="kanban-card-editor">
                      <div className="kanban-editor-heading">
                        <p className="eyebrow">Card selecionado</p>
                        <h4>{selectedKanbanIssue ? selectedKanbanIssue.title : 'Selecione um card para editar'}</h4>
                        <span>Arraste o card entre os marcos ou ajuste as informações principais abaixo.</span>
                      </div>

                      {selectedKanbanIssue ? (
                        <form className="kanban-editor-form" onSubmit={saveKanbanCardEdit}>
                          <label className="wide-field">
                            Título do issue
                            <input value={kanbanEditDraft.title} onChange={(event) => updateKanbanEditDraft('title', event.target.value)} required />
                          </label>
                          <label>
                            Marco Contratual
                            <select value={kanbanEditDraft.marcoId} onChange={(event) => updateKanbanEditDraft('marcoId', event.target.value)}>
                              <option value="">Sem marco vinculado</option>
                              {linkKanbanColumns.filter((column) => column.canReceiveCards).map((column) => (
                                <option key={column.id} value={column.id}>{getKanbanColumnOptionLabel(column)}</option>
                              ))}
                            </select>
                          </label>
                          <label>
                            Atribuir
                            <select value={kanbanEditDraft.assignedTo} onChange={(event) => updateKanbanEditDraft('assignedTo', event.target.value)}>
                              <option value="">Sem responsável</option>
                              {projectUsers.map((projectUser) => (
                                <option key={projectUser.id} value={projectUser.id}>
                                  {projectUser.name}{projectUser.email ? ` - ${projectUser.email}` : ''}
                                </option>
                              ))}
                            </select>
                          </label>
                          <label className="wide-field">
                            Seguidores
                            <select
                              multiple
                              className="multi-select"
                              value={kanbanEditDraft.followers}
                              onChange={(event) =>
                                updateKanbanEditDraft(
                                  'followers',
                                  Array.from(event.target.selectedOptions, (option) => option.value)
                                )
                              }
                            >
                              {projectUsers.map((projectUser) => (
                                <option key={projectUser.id} value={projectUser.id}>
                                  {projectUser.name}{projectUser.email ? ` - ${projectUser.email}` : ''}
                                </option>
                              ))}
                            </select>
                            <small>Segure Ctrl para selecionar mais de uma pessoa.</small>
                          </label>
                          <label>
                            Data prevista
                            <input type="date" value={kanbanEditDraft.dueDate} onChange={(event) => updateKanbanEditDraft('dueDate', event.target.value)} />
                          </label>
                          <label>
                            Disciplina
                            <select value={kanbanEditDraft.disciplina} onChange={(event) => updateKanbanEditDraft('disciplina', event.target.value)}>
                              <option value="">Não informado</option>
                              {kanbanDisciplineOptions.map((option) => (
                                <option key={option.id} value={option.id}>{option.label}</option>
                              ))}
                            </select>
                          </label>
                          <label className="wide-field">
                            Descrição
                            <textarea value={kanbanEditDraft.description} onChange={(event) => updateKanbanEditDraft('description', event.target.value)} />
                          </label>
                          <label className="wide-field">
                            Comentário / observação
                            <textarea
                              value={kanbanEditDraft.comment}
                              onChange={(event) => updateKanbanEditDraft('comment', event.target.value)}
                              placeholder="Opcional. Nesta etapa o comentário fica registrado na descrição do issue."
                            />
                          </label>
                          <div className="kanban-editor-actions">
                            <button type="button" className="secondary-button" onClick={() => setSelectedIssueId(selectedKanbanIssue.id)}>
                              Ver ficha completa
                            </button>
                            <button type="submit" className="primary-button" disabled={kanbanSavingIssueId === selectedKanbanIssue.id || savingIssueId === selectedKanbanIssue.id}>
                              {kanbanSavingIssueId === selectedKanbanIssue.id || savingIssueId === selectedKanbanIssue.id ? 'Salvando...' : 'Salvar no ACC'}
                            </button>
                          </div>
                        </form>
                      ) : (
                        <p className="empty-state">Clique em um card do quadro para editar responsável, data, disciplina, descrição e vínculo com o marco.</p>
                      )}
                    </section>
                  </>
                )}

                {false && issueLinkGroups.length > 0 && (
                  <div className="flow-diagram" aria-label="Issues do ACC agrupados por Pacote/Marco Contratual e conectados pelas Referências do ACC">
                    <div className="flow-color-legend" aria-label="Legenda do mapa de vínculos">
                      <span><i className="legend-milestone" /> Marco de entrega</span>
                      <span><i className="legend-dependency" /> Dependência vinculada</span>
                      <span><i className="legend-done" /> Concluída</span>
                      <span><i className="legend-late" /> Atrasada</span>
                    </div>
                    {issueLinkGroups.map(({ category, milestone, dependencies, flows, total, linked, open, closed, overdue, progress, isEmpty }, categoryIndex) => (
                      <section key={category} className={`flow-lane flow-lane-${categoryIndex % 4}`}>
                        <div className="flow-lane-label">
                          <strong>{category}</strong>
                        </div>
                        <div className="flow-lane-canvas">
                          <div className="flow-lane-metrics">
                            <span>{progress}% Avanço</span>
                            <span>{overdue} Atrasados</span>
                            <span>{closed} Concluídos</span>
                            <span>{total} Total</span>
                            <span>{linked} Issues vinculados</span>
                          </div>
                          {milestone ? (
                            (dependencies.length ? dependencies : [null]).map((dependency, dependencyIndex) => (
                              <article key={`${milestone.id}-${dependency?.id || 'sem-dependencia'}-${dependencyIndex}`} className={`flow-row ${dependency ? 'has-links' : ''}`}>
                                {dependencyIndex === 0 ? (
                                  <button
                                    type="button"
                                    className={`flow-node flow-node-source flow-node-milestone ${!isOpenIssue(milestone) ? 'is-closed' : ''} ${isOverdue(milestone) ? 'is-overdue' : ''}`}
                                    onClick={() => setSelectedIssueId(milestone.id)}
                                    style={{ '--discipline-color': getDisciplineColor(getIssueDiscipline(milestone)) }}
                                  >
                                    <strong>{milestone.displayId || milestone.id} - {milestone.title}</strong>
                                    <small>Marco de entrega | {formatDate(milestone.startDate || milestone.createdAt)} → {formatDate(milestone.dueDate || milestone.updatedAt)}</small>
                                  </button>
                                ) : (
                                  <span className="flow-node-spacer" aria-hidden="true" />
                                )}
                                {dependency ? (
                                  <>
                                    <span className="flow-line" aria-hidden="true" />
                                    <button
                                      type="button"
                                      className={`flow-node flow-node-target ${!isOpenIssue(dependency) ? 'is-closed' : ''} ${isOverdue(dependency) ? 'is-overdue' : ''}`}
                                      onClick={() => setSelectedIssueId(dependency.id)}
                                      style={{ '--discipline-color': getDisciplineColor(getIssueDiscipline(dependency)) }}
                                    >
                                      <strong>{dependency.displayId || dependency.id} - {dependency.title}</strong>
                                      <small>{dependency.issueType || 'Sem tipo'} | {formatDate(dependency.startDate || dependency.createdAt)} → {formatDate(dependency.dueDate || dependency.updatedAt)} | {getIssueDiscipline(dependency)}</small>
                                    </button>
                                  </>
                                ) : (
                                  <div className="flow-node flow-node-empty">Nenhuma dependência vinculada a este marco.</div>
                                )}
                              </article>
                            ))
                          ) : (
                            flows.map((issue) => (
                              <article key={issue.id} className="flow-row">
                                <button
                                  type="button"
                                  className={`flow-node flow-node-source ${!isOpenIssue(issue) ? 'is-closed' : ''} ${isOverdue(issue) ? 'is-overdue' : ''}`}
                                  onClick={() => setSelectedIssueId(issue.id)}
                                  style={{ '--discipline-color': getDisciplineColor(getIssueDiscipline(issue)) }}
                                >
                                  <strong>{issue.displayId || issue.id} - {issue.title}</strong>
                                  <small>{issue.issueType || 'Sem tipo'} | {formatDate(issue.startDate || issue.createdAt)} → {formatDate(issue.dueDate || issue.updatedAt)} | {getIssueDiscipline(issue)}</small>
                                </button>
                              </article>
                            ))
                          )}
                          {isEmpty && milestone && <p className="flow-empty-note">Use o campo Referências do ACC ou o Código do Marco para vincular dependências.</p>}
                        </div>
                      </section>
                    ))}
                  </div>
                )}

                {selectedIssue && (
                  <section className="issue-detail-panel compact-detail">
                    <div>
                      <p className="eyebrow">Issue selecionada</p>
                      <h3>{selectedIssue.title}</h3>
                      <p className="muted">ID Autodesk: {selectedIssue.id}</p>
                    </div>
                    <div className="detail-grid">
                      <span>Status: <strong>{selectedIssue.status}</strong></span>
                      <span>Prazo: <strong>{formatDate(selectedIssue.dueDate)}</strong></span>
                      <span>Categoria: <strong>{selectedIssue.category || 'Sem categoria'}</strong></span>
                      <span>Tipo: <strong>{selectedIssue.issueType || 'Sem tipo'}</strong></span>
                    </div>
                  </section>
                )}
              </section>
            )}

            {activeModule === 'interfaces-pendencias' && (
              <section className="interfaces-panel">
                <div className="interfaces-summary-card">
                  <div>
                    <p className="eyebrow">Painel executivo</p>
                    <h3>INTERFACES E PENDÊNCIAS</h3>
                    <p className="muted">Leitura dos issues, revisões e pacotes emitidos para apoiar a gestão de interfaces.</p>
                  </div>
                  <div className="interfaces-kpis">
                    <span><strong>{interfacesPendenciasDashboard.total}</strong> issues</span>
                    <span><strong>{interfacesPendenciasDashboard.open}</strong> abertas</span>
                    <span><strong>{interfacesPendenciasDashboard.overdue}</strong> atrasadas</span>
                    <span><strong>{interfacesPendenciasDashboard.encaminhamentos}</strong> encaminhamentos</span>
                  </div>
                </div>

                <div className="interfaces-package-control">
                  <div className="interfaces-card-heading">
                    <div>
                      <h4>Revisões, pacotes e consumo técnico</h4>
                      <span>Baseado na lista mestre e nos arquivos publicados disponíveis ao usuário logado.</span>
                    </div>
                    <button className="ghost-button" type="button" onClick={() => loadPublishedDocuments(selectedProjectId, { forceRefresh: true })} disabled={documentsLoading}>
                      {documentsLoading ? 'Atualizando...' : 'Atualizar documentos'}
                    </button>
                  </div>
                  <div className="interfaces-doc-kpis">
                    <span><strong>{interfacesPackageDashboard.total}</strong> documentos na lista</span>
                    <span><strong>{interfacesPackageDashboard.emitted}</strong> emitidos</span>
                    <span><strong>{interfacesPackageDashboard.pending}</strong> pendentes</span>
                    <span><strong>{interfacesPackageDashboard.packageCount}</strong> GRDs / pacotes</span>
                  </div>
                  {!interfacesPackageDashboard.hasData ? (
                    <p className="muted">Ainda não carreguei a lista de documentos deste projeto. Clique em Atualizar documentos para cruzar revisões, GRDs e arquivos publicados.</p>
                  ) : (
                    <div className="interfaces-package-grid">
                      <article>
                        <h5>Últimas revisões emitidas</h5>
                        <div className="interfaces-revision-list">
                          {interfacesPackageDashboard.revisions.map((document) => (
                            <a key={`${document.id}:${document.revision}:${document.version}`} href={document.webView || '#'} target="_blank" rel="noreferrer" className={!document.webView ? 'disabled-link' : ''}>
                              <strong>{document.code}</strong>
                              <span>{document.discipline} • Rev. {document.revision} • {document.version} • {document.date}</span>
                            </a>
                          ))}
                          {interfacesPackageDashboard.revisions.length === 0 && <span className="muted">Nenhuma revisão identificada nos documentos emitidos.</span>}
                        </div>
                      </article>
                      <article>
                        <h5>Pacotes / GRDs recentes</h5>
                        <div className="interfaces-package-list">
                          {interfacesPackageDashboard.packages.map((pack) => (
                            <div key={pack.key}>
                              <strong>{pack.key}</strong>
                              <span>{pack.documents} docs • {pack.disciplineCount} disciplinas • {pack.latestDateLabel}</span>
                            </div>
                          ))}
                          {interfacesPackageDashboard.packages.length === 0 && <span className="muted">Nenhum pacote ou GRD identificado.</span>}
                        </div>
                      </article>
                      <article>
                        <h5>Consumo por disciplina</h5>
                        <div className="interfaces-consumption-list">
                          {interfacesPackageDashboard.disciplines.map((discipline) => (
                            <div key={discipline.discipline}>
                              <span><strong>{discipline.discipline}</strong><b>{discipline.percent}%</b></span>
                              <i><em style={{ width: `${discipline.percent}%` }} /></i>
                              <small>{discipline.emitted} emitidos • {discipline.pending} pendentes • {discipline.packageCount} pacotes</small>
                            </div>
                          ))}
                        </div>
                      </article>
                    </div>
                  )}
                  <p className="interfaces-api-note">Observação: o app já controla revisões e pacotes publicados. A confirmação formal de “consumo” do Design Collaboration depende da Autodesk disponibilizar essa leitura no projeto/API; enquanto isso, esta visão usa emissão, GRD e pendências por disciplina como indicador gerencial.</p>
                </div>

                <div className="interfaces-filter-bar">
                  <label>
                    Disciplina
                    <select value={interfacesDisciplineFilter} onChange={(event) => setInterfacesDisciplineFilter(event.target.value)}>
                      <option value="all">Todas</option>
                      {interfacesPendenciasDashboard.disciplineOptions.map((discipline) => (
                        <option key={discipline} value={discipline}>{discipline}</option>
                      ))}
                    </select>
                  </label>
                  <label>
                    Status
                    <select value={interfacesStatusFilter} onChange={(event) => setInterfacesStatusFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      <option value="open">Abertas</option>
                      <option value="overdue">Atrasadas</option>
                      <option value="closed">Fechadas</option>
                    </select>
                  </label>
                  <label>
                    Impacto
                    <select value={interfacesImpactFilter} onChange={(event) => setInterfacesImpactFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      <option value="Alto">Alto</option>
                      <option value="Médio">Médio</option>
                      <option value="Baixo">Baixo</option>
                      <option value="Sem classificação">Sem classificação</option>
                    </select>
                  </label>
                  <button
                    type="button"
                    className="ghost-button"
                    onClick={() => {
                      setInterfacesDisciplineFilter('all');
                      setInterfacesStatusFilter('all');
                      setInterfacesImpactFilter('all');
                    }}
                  >
                    Limpar filtros
                  </button>
                </div>

                <div className="interfaces-grid">
                  <article className="interfaces-card interfaces-card-large">
                    <div className="interfaces-card-heading">
                      <h4>Pendências por disciplina</h4>
                      <span>{interfacesPendenciasDashboard.disciplines.length} disciplinas</span>
                    </div>
                    <div className="discipline-list">
                      {interfacesPendenciasDashboard.disciplines.map((discipline) => (
                        <button
                          key={discipline.discipline}
                          type="button"
                          className="discipline-row"
                          onClick={() => discipline.issues[0] && setSelectedIssueId(discipline.issues[0].id)}
                        >
                          <i style={{ background: getDisciplineColor(discipline.discipline) }} />
                          <span>
                            <strong>{discipline.discipline}</strong>
                            <small>{discipline.open} abertas • {discipline.overdue} atrasadas • {discipline.highImpact} alto impacto</small>
                          </span>
                          <b>{discipline.total}</b>
                        </button>
                      ))}
                    </div>
                    {interfacesPendenciasDashboard.disciplines.length === 0 && (
                      <p className="muted">Nenhuma pendência encontrada para os filtros selecionados.</p>
                    )}
                  </article>

                  <article className="interfaces-card">
                    <div className="interfaces-card-heading">
                      <h4>Impactos</h4>
                      <span>Prazo, escopo, medição e qualidade</span>
                    </div>
                    <div className="impact-mini-grid">
                      {interfacesPendenciasDashboard.impactSummary.map((impact) => (
                        <div key={impact.key} className="impact-mini-card">
                          <strong>{impact.label}</strong>
                          <span className="impact-mini-line high"><em />Alto <b>{impact.counts.Alto || 0}</b></span>
                          <span className="impact-mini-line medium"><em />Médio <b>{impact.counts.Médio || 0}</b></span>
                          <span className="impact-mini-line low"><em />Baixo <b>{impact.counts.Baixo || 0}</b></span>
                        </div>
                      ))}
                    </div>
                  </article>

                  <article className="interfaces-card">
                    <div className="interfaces-card-heading">
                      <h4>Reuniões e encaminhamentos</h4>
                      <span>Itens com contexto de ata, reunião ou responsável</span>
                    </div>
                    <div className="routing-summary">
                      <div><strong>{interfacesPendenciasDashboard.meetings}</strong><span>reuniões / atas identificadas</span></div>
                      <div><strong>{interfacesPendenciasDashboard.encaminhamentos}</strong><span>encaminhamentos em aberto</span></div>
                    </div>
                    <div className="interface-highlight-list">
                      {interfacesPendenciasDashboard.highlightedIssues.map((issue) => (
                        <button key={issue.id} type="button" onClick={() => setSelectedIssueId(issue.id)}>
                          <span>{issue.displayId || issue.id}</span>
                          <strong>{issue.title}</strong>
                          <small>{getIssueDiscipline(issue)} • {formatDate(issue.dueDate || issue.createdAt)}</small>
                        </button>
                      ))}
                    </div>
                  </article>
                </div>

                {selectedIssueFull && (
                  <section className="issue-detail-panel compact-detail">
                    <div>
                      <p className="eyebrow">Issue selecionada</p>
                      <h3>{selectedIssueFull.title}</h3>
                      <p className="muted">ID Autodesk: {selectedIssueFull.id}</p>
                    </div>
                    <div className="detail-grid">
                      <span>Status: <strong>{selectedIssueFull.status || 'Não informado'}</strong></span>
                      <span>Prazo: <strong>{formatDate(selectedIssueFull.dueDate)}</strong></span>
                      <span>Disciplina: <strong>{getIssueDiscipline(selectedIssueFull)}</strong></span>
                      <span>Responsável: <strong>{formatDetailValue(selectedIssueFull.assignedTo || selectedIssueFull.assignee || selectedIssueFull.raw?.assignedTo)}</strong></span>
                      <span>Tipo: <strong>{selectedIssueFull.issueType || 'Sem tipo'}</strong></span>
                      <span>Categoria: <strong>{selectedIssueFull.category || 'Sem categoria'}</strong></span>
                      <span>Criada em: <strong>{formatDateTime(selectedIssueFull.createdAt)}</strong></span>
                      <span>Atualizada em: <strong>{formatDateTime(selectedIssueFull.updatedAt)}</strong></span>
                    </div>
                    {selectedIssueFull.description && (
                      <p className="issue-detail-description">{selectedIssueFull.description}</p>
                    )}
                  </section>
                )}
              </section>
            )}

            {activeModule === 'meetings' && (
              <section className="meetings-panel">
                <div className="meetings-grid">
                  {meetingTemplates.map((template) => (
                    <button
                      key={template.id}
                      className={`meeting-template-card ${selectedMeetingTemplateId === template.id ? 'selected' : ''}`}
                      type="button"
                      onClick={() => selectMeetingTemplate(template.id)}
                    >
                      <span className="module-icon" aria-hidden="true"><UiIcon name="meeting" /></span>
                      <strong>{template.label}</strong>
                      <span>{template.purpose}</span>
                    </button>
                  ))}
                </div>

                <article className="meeting-draft-card">
                  <div className="meeting-draft-heading">
                    <div>
                      <p className="eyebrow">Modelo selecionado</p>
                      <h3>{selectedMeetingTemplate.title}</h3>
                      <p>{selectedMeetingTemplate.purpose}</p>
                    </div>
                    <button className="ghost-button" type="button" onClick={exportMeetingDraftWord}>
                      Exportar Word
                    </button>
                  </div>

                  <div className="meeting-form-grid">
                    <label>
                      Data da reuniao
                      <input type="date" value={meetingDate} onChange={(event) => setMeetingDate(event.target.value)} />
                    </label>
                    <label>
                      Participantes
                      <input
                        type="text"
                        value={meetingParticipants}
                        onChange={(event) => setMeetingParticipants(event.target.value)}
                        placeholder="Ex.: G5, cliente, disciplinas envolvidas"
                      />
                    </label>
                  </div>

                  <label className="meeting-topics-field">
                    Pauta da reuniao
                    <textarea
                      value={meetingCustomTopics}
                      onChange={(event) => setMeetingCustomTopics(event.target.value)}
                      rows={8}
                    />
                  </label>

                  <div className="meeting-output">
                    <h4>Roteiro para preencher no ACC</h4>
                    <p><strong>Projeto:</strong> {selectedProject?.name || selectedProjectId}</p>
                    <p><strong>Data:</strong> {meetingDate ? formatDate(meetingDate) : 'A definir'}</p>
                    <p><strong>Participantes:</strong> {meetingParticipants || 'A definir'}</p>
                    <ol>
                      {meetingCustomTopics
                        .split('\n')
                        .map((topic) => topic.trim())
                        .filter(Boolean)
                        .map((topic, index) => (
                          <li key={`${topic}-${index}`}>{topic}</li>
                        ))}
                    </ol>
                  </div>

                  <div className="meeting-api-note">
                    <strong>Integracao ACC</strong>
                    <span>Este modulo prepara o roteiro. A criacao direta em Meetings/Design Collaboration depende de endpoint publico da Autodesk para reunioes.</span>
                  </div>
                </article>
              </section>
            )}

            {activeModule === 'documents' && (
              <section className="documents-panel documents-panel-pro">
                <div className="documents-module-header">
                  <div>
                    <p className="eyebrow">Documentos</p>
                    <h2>DOCUMENTOS EMITIDOS</h2>
                    <span>
                      Projeto selecionado: <strong>{selectedProject?.name || selectedProjectId}</strong>
                    </span>
                  </div>
                  <div className="documents-header-actions">
                    <button className="icon-button-light" type="button" onClick={() => setActiveModule('')}>
                      <UiIcon name="modules" /> Modulos
                    </button>
                    <button className="icon-button-light" type="button" onClick={exportPublishedDocumentsCsv} disabled={!documentListRows.length}>
                      <UiIcon name="export" /> Exportar planilha
                    </button>
                    <button className="icon-button-light" type="button" onClick={exportPublishedDocumentsPowerBiCsv} disabled={!documentListRows.length}>
                      <UiIcon name="export" /> Power BI
                    </button>
                    <button className="icon-button-primary" type="button" onClick={() => loadPublishedDocuments(selectedProjectId, { forceRefresh: true })} disabled={documentsLoading || documentsSpreadsheetSaving}>
                      <UiIcon name="refresh" /> Atualizar dados
                    </button>
                    <button
                      className="icon-button-primary"
                      type="button"
                      onClick={updatePublishedDocumentsSpreadsheet}
                      disabled={!documentListRows.length || documentsLoading || documentsSpreadsheetSaving}
                    >
                      <UiIcon name="export" /> {documentsSpreadsheetSaving ? 'Salvando...' : 'Atualizar planilha no ACC'}
                    </button>
                    {documentsUpdatedAt && <span className="documents-updated">Ultima atualizacao: {formatDateTime(documentsUpdatedAt)}</span>}
                  </div>
                </div>

                <div className="documents-master-card">
                  <div className="documents-master-block">
                    <span className="documents-card-icon"><UiIcon name="file" /></span>
                    <div>
                      <p className="eyebrow">Lista mestre</p>
                      <h3>DOCUMENTOS EMITIDOS</h3>
                      <span className="muted">
                        {documentListSpreadsheet
                          ? documentListSpreadsheet.name
                          : publishedDocumentsMessage || 'Buscando a planilha em 05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS.'}
                      </span>
                      {documentListSpreadsheet && <span className="muted">{documentListSpreadsheet.path}</span>}
                    </div>
                  </div>
                  <div className="documents-master-block">
                    <span className="documents-card-icon folder"><UiIcon name="folder" /></span>
                    <div>
                      <p>Base de comparacao</p>
                      <strong>{formatDocumentFolderLabel(publishedFolder?.path || 'Project Files / 04. PUBLICADOS (DocEmit)')}</strong>
                      <span className="muted">Pasta do CDE utilizada para verificar os documentos publicados.</span>
                    </div>
                  </div>
                  <div className="documents-master-block summary">
                    <p>Resumo do projeto</p>
                    <span>Disciplinas envolvidas <strong>{documentSummary.disciplines}</strong></span>
                    <span>Tipos de documento <strong>{documentSummary.documentTypes}</strong></span>
                    <span>Pastas com emissao <strong>{documentSummary.folders}</strong></span>
                  </div>
                </div>

                <div className="document-metrics">
                  <div className="metric-card metric-blue">
                    <span className="metric-icon"><UiIcon name="file" /></span>
                    <span>Total na lista</span>
                    <strong>{documentMetrics.total}</strong>
                    <small>Documentos previstos</small>
                  </div>
                  <div className="metric-card metric-green">
                    <span className="metric-icon"><UiIcon name="check" /></span>
                    <span>Emitidos</span>
                    <strong>{documentMetrics.emitted}</strong>
                    <small>Documentos publicados</small>
                  </div>
                  <div className="metric-card metric-percentage">
                    <span className="metric-icon"><UiIcon name="overview" /></span>
                    <span>Percentual do projeto</span>
                    <span className="metric-highlight">{documentEmittedPercentage}%</span>
                    <small>Ja emitido</small>
                  </div>
                  <div className="metric-card metric-yellow">
                    <span className="metric-icon"><UiIcon name="clock" /></span>
                    <span>Pendentes</span>
                    <strong>{documentMetrics.pending}</strong>
                    <small>Documentos nao publicados</small>
                  </div>
                  <div className="metric-card metric-purple">
                    <span className="metric-icon"><UiIcon name="folder" /></span>
                    <span>Com GRD</span>
                    <strong>{documentMetrics.withGrd}</strong>
                    <small>Documentos com GRD</small>
                  </div>
                </div>

                <div className="documents-filter-bar">
                  <label className="documents-search">
                    <UiIcon name="search" />
                    <input
                      value={documentSearch}
                      onChange={(event) => setDocumentSearch(event.target.value)}
                      placeholder="Buscar por codigo, descricao ou nome do arquivo..."
                    />
                  </label>
                  <label>
                    <span>Status</span>
                    <select value={documentStatusFilter} onChange={(event) => setDocumentStatusFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      <option value="emitted">Emitidos</option>
                      <option value="pending">Pendentes</option>
                    </select>
                  </label>
                  <label>
                    <span>Tipo de documento</span>
                    <select value={documentTypeFilter} onChange={(event) => setDocumentTypeFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {documentTypeOptions.map((type) => (
                        <option key={type} value={type}>
                          {type}
                        </option>
                      ))}
                    </select>
                  </label>
                  <button className="icon-button-light" type="button" onClick={clearDocumentFilters}>
                    <UiIcon name="close" /> Limpar filtros
                  </button>
                </div>

                {documentsLoading && (
                  <div className="loading-box">
                    <strong>Montando comparativo...</strong>
                    <span>Estou lendo a planilha da pasta 05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS e comparando com 04. PUBLICADOS (DocEmit) e subpastas.</span>
                  </div>
                )}

                {publishedDocumentsPartial && (
                  <div className="loading-box">
                    <strong>Comparativo parcial carregado.</strong>
                    <span>Usei os primeiros documentos publicados encontrados para evitar limite de quota da Autodesk.</span>
                  </div>
                )}

                {documentsSpreadsheetSaveMessage && (
                  <div className="loading-box">
                    <strong>Atualizacao da planilha</strong>
                    <span>{documentsSpreadsheetSaveMessage}</span>
                  </div>
                )}

                {!documentsLoading && documentListRows.length === 0 && (
                  <div className="empty-state">
                    <strong>Nao encontrei a lista de documentos para comparar.</strong>
                    <span>
                      {publishedDocumentsMessage ||
                        'Confirme se existe uma planilha Excel na pasta 05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS e se existe a pasta 04. PUBLICADOS (DocEmit) no projeto.'}
                    </span>
                  </div>
                )}

                {documentListRows.length > 0 && (
                  <div className="documents-table-shell">
                    <div className="table-scrollbar-top" ref={documentTopScrollRef} onScroll={() => syncDocumentTableScroll('top')}>
                      <div className="table-scrollbar-spacer" />
                    </div>
                    <div
                      className="documents-table-wrap documents-table-wrap-pro"
                      ref={documentTableScrollRef}
                      onScroll={() => syncDocumentTableScroll('table')}
                    >
                    <table className="documents-table document-register-table">
                      <thead>
                        <tr>
                          <th>
                            <button className="table-sort-button" type="button" onClick={() => changeDocumentSort('discipline')}>
                              Disciplina {documentSort.field === 'discipline' ? (documentSort.direction === 'asc' ? '↑' : '↓') : ''}
                            </button>
                          </th>
                          <th>
                            <button className="table-sort-button" type="button" onClick={() => changeDocumentSort('code')}>
                              Codigo de engenharia {documentSort.field === 'code' ? (documentSort.direction === 'asc' ? '↑' : '↓') : ''}
                            </button>
                          </th>
                          <th>Descricao</th>
                          <th>Tipo de documento</th>
                          <th>Nome</th>
                          <th>Rev.</th>
                          <th>Versao</th>
                          <th>GRD</th>
                          <th>Status</th>
                          <th>Data de emissao</th>
                        </tr>
                      </thead>
                      <tbody>
                        {sortedDocumentRows.map((document) => (
                          <tr key={document.id} className={document.emitted ? 'document-emitted-row' : 'document-pending-row'}>
                            <td>{document.discipline || '-'}</td>
                            <td>{document.code || '-'}</td>
                            <td>{document.title || '-'}</td>
                            <td>{document.documentType || '-'}</td>
                            <td>
                              {document.webView ? (
                                <a href={document.webView} target="_blank" rel="noreferrer">
                                  {document.emittedFileName}
                                </a>
                              ) : (
                                document.emittedFileName || '-'
                              )}
                            </td>
                            <td>{document.emittedRevision || document.revision || '-'}</td>
                            <td>{document.emittedVersion || '-'}</td>
                            <td>{document.emittedGrd || document.grd || '-'}</td>
                            <td>
                              <span className={document.emitted ? 'document-status emitted' : 'document-status pending'}>
                                {document.status || (document.emitted ? 'Emitido' : 'Pendente')}
                              </span>
                            </td>
                            <td>{document.emittedDate || document.emissionDate ? formatDate(document.emittedDate || document.emissionDate) : '-'}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                    </div>
                    <div className="documents-table-footer">
                      <span>
                        Mostrando {sortedDocumentRows.length} de {documentListRows.length} documentos
                      </span>
                    </div>
                  </div>
                )}

                {false && documentsByFolder.length > 0 && (
                  <div className="documents-folder-list">
                    {documentsByFolder.map((group) => (
                      <section key={group.folderName} className="documents-folder-group">
                        <div className="documents-folder-header">
                          <h3>{group.folderName}</h3>
                          <span>{group.documents.length} documentos</span>
                        </div>
                        <div className="documents-table-wrap">
                          <table className="documents-table">
                            <thead>
                              <tr>
                                <th>Arquivo</th>
                                <th>GRD</th>
                                <th>Descrição</th>
                                <th>Rev</th>
                                <th>Versão</th>
                                <th>Data da versão</th>
                                <th>Atualizado em</th>
                              </tr>
                            </thead>
                            <tbody>
                              {group.documents.map((document) => (
                                <tr key={document.id}>
                                  <td>
                                    {document.webView ? (
                                      <a href={document.webView} target="_blank" rel="noreferrer">
                                        {document.name}
                                      </a>
                                    ) : (
                                      document.name
                                    )}
                                  </td>
                                  <td>{document.grd || '-'}</td>
                                  <td>{document.description || '-'}</td>
                                  <td>{document.revision || '-'}</td>
                                  <td>{document.version || '-'}</td>
                                  <td>{document.versionCreatedAt ? formatDate(document.versionCreatedAt) : '-'}</td>
                                  <td>{document.updatedAt ? formatDate(document.updatedAt) : '-'}</td>
                                </tr>
                              ))}
                            </tbody>
                          </table>
                        </div>
                      </section>
                    ))}
                  </div>
                )}
              </section>
            )}

            {activeModule === 'im' && (
              <section className="eap-panel">
                {!imMode && (
                  <div className="module-grid im-module-grid">
                    <button className="module-button module-button-primary" type="button" onClick={() => { setImMode('worksheet'); resetImModule(); }}>
                      <span className="module-copy"><strong>Planilha de Issues</strong><span>Visualizar, exportar e editar em massa os issues do ACC em formato de planilha, incluindo campos personalizados.</span></span>
                    </button>
                    <button className="module-button module-button-links" type="button" onClick={() => { setImMode('update'); resetImModule(); }}>
                      <span className="module-copy"><strong>Atualizar Issues</strong><span>Atualizar issues existentes com prévia comparativa.</span></span>
                    </button>
                  </div>
                )}
                {imMode === 'worksheet' && (
                  <>
                    <div className="eap-intro">
                      <div>
                        <p className="eyebrow">Planilha de Issues</p>
                        <h3>Planilha de Issues</h3>
                        <span>Carregue os issues do projeto ACC em formato de planilha para editar datas, status, impactos e campos personalizados em massa.</span>
                      </div>
                      <div className="section-actions">
                        <button className="ghost-button" type="button" onClick={loadImWorksheet} disabled={imLoading}>Carregar issues do ACC</button>
                        <button className="ghost-button" type="button" onClick={() => {
                          const wb = XLSX.utils.book_new();
                          XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(imSheetRows), 'PlanilhaIssues');
                          XLSX.writeFile(wb, `planilha-issues-${new Date().toISOString().slice(0,10)}.xlsx`);
                        }} disabled={!imSheetRows.length}>Exportar issues para Excel</button>
                      </div>
                    </div>
                    {!!imSheetRows.length && (
                      <>
                        <p>Issues carregados: <strong>{imSheetRows.length}</strong> | Linhas alteradas: <strong>{new Set(Object.keys(imSheetChanges).map((key) => key.split(':')[0])).size}</strong> | Células alteradas: <strong>{Object.keys(imSheetChanges).length}</strong></p>
                        <div className="documents-table-wrap eap-table-wrap">
                          <table className="documents-table eap-preview-table">
                            <thead><tr><th>Selecionar</th><th>Issue ID</th><th>Título</th><th>Status</th><th>Responsável</th><th>Data de vencimento</th>{imSheetCustomColumns.map((col) => <th key={col.id}>{col.name}</th>)}</tr></thead>
                            <tbody>{imSheetRows.slice(0, 300).map((row, rowIndex) => <tr key={row.accIssueId || row.issueId || rowIndex}>
                              <td><input type="checkbox" checked={Boolean(imSheetSelectedRows[row.accIssueId || row.issueId])} onChange={(e)=>setImSheetSelectedRows((prev)=>({...prev,[row.accIssueId || row.issueId]: e.target.checked}))} /></td>
                              <td>{row.issueId}</td>
                              <td><input value={row.title} onChange={(e)=>updateWorksheetCell(rowIndex, 'title', e.target.value)} /></td>
                              <td><select value={row.status} onChange={(e)=>updateWorksheetCell(rowIndex, 'status', e.target.value)}><option value="">Selecione</option>{statusOptions.map((status) => <option key={status.value} value={status.value}>{status.label}</option>)}</select></td>
                              <td><input value={row.assignedTo} onChange={(e)=>updateWorksheetCell(rowIndex, 'assignedTo', e.target.value)} /></td>
                              <td><input type="date" value={row.dueDate} onChange={(e)=>updateWorksheetCell(rowIndex, 'dueDate', e.target.value)} /></td>
                              {imSheetCustomColumns.map((col) => <td key={`${row.accIssueId || row.issueId}-${col.id}`}>{col.type === 'list' ? <select value={row[`cf_${col.id}`] || ''} onChange={(e)=>updateWorksheetCell(rowIndex, `cf_${col.id}`, e.target.value)}><option value="">Selecione</option>{col.options.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}</select> : <input type={col.type === 'date' ? 'date' : col.type === 'number' ? 'number' : 'text'} value={row[`cf_${col.id}`] || ''} onChange={(e)=>updateWorksheetCell(rowIndex, `cf_${col.id}`, e.target.value)} />}</td>)}
                            </tr>)}</tbody>
                          </table>
                        </div>
                        <div className="section-actions">
                          <button className="ghost-button" type="button" onClick={() => {
                            const previewRows = [];
                            imSheetRows.forEach((row, index) => {
                              const original = imSheetOriginalRows[index] || {};
                              Object.keys(row).forEach((field) => {
                                if (row[field] !== original[field]) previewRows.push({ issueId: row.issueId, title: row.title, field, oldValue: original[field] || '', newValue: row[field] || '', validation: 'Pronto' });
                              });
                            });
                            setImSheetPreviewRows(previewRows);
                          }}>Pré-visualizar alterações</button>
                          <button className="primary-button" type="button" onClick={async () => {
                            setImLoading(true);
                            setError('');
                            setImWorksheetFeedback({ type: 'info', message: 'Atualizando issues no ACC... aguarde.' });
                            try {
                              const validationErrors = [];
                              const pendingUpdates = [];
                              const results = [];
                              for (const [idx, row] of imSheetRows.entries()) {
                                const original = imSheetOriginalRows[idx] || {};
                                const changes = {};
                                ['title', 'status', 'dueDate'].forEach((field) => {
                                  if (row[field] !== original[field]) changes[field] = row[field] || null;
                                });
                                const changedCustomAttributes = imSheetCustomColumns
                                  .filter((col) => row[`cf_${col.id}`] !== original[`cf_${col.id}`])
                                  .map((col) => ({ id: col.id, value: row[`cf_${col.id}`] || null, name: col.name, type: col.type, options: col.options, required: col.required }));
                                if (!row.accIssueId) validationErrors.push(`Linha ${idx + 1}: issue sem ID ACC válido.`);
                                if (row.dueDate && Number.isNaN(new Date(row.dueDate).getTime())) validationErrors.push(`Linha ${idx + 1}: data inválida.`);
                                changedCustomAttributes.forEach((attribute) => {
                                  if (!attribute.id) validationErrors.push(`Linha ${idx + 1}: campo personalizado sem mapeamento ACC.`);
                                  if (attribute.required && !attribute.value) validationErrors.push(`Linha ${idx + 1}: campo obrigatório "${attribute.name}" sem valor.`);
                                  if (attribute.type === 'list' && attribute.value && !attribute.options.some((option) => option.value === attribute.value)) {
                                    validationErrors.push(`Linha ${idx + 1}: valor inválido para "${attribute.name}".`);
                                  }
                                });
                                if (changedCustomAttributes.length) changes.customAttributes = changedCustomAttributes.map((attribute) => ({ id: attribute.id, value: attribute.value }));
                                if (Object.keys(changes).length) pendingUpdates.push({ row, changes });
                              }
                              if (validationErrors.length) {
                                setError(`Erros de validação:\n${validationErrors.join('\n')}`);
                                setImWorksheetFeedback({ type: 'error', message: 'Foram encontrados erros de validação. Revise os campos e tente novamente.' });
                                return;
                              }
                              for (const item of pendingUpdates) {
                                try {
                                  await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/issues/${encodeURIComponent(item.row.accIssueId)}`, { method: 'PATCH', body: JSON.stringify(item.changes) });
                                  results.push({ issueId: item.row.issueId, status: 'Sucesso', message: 'Atualizado' });
                                } catch (e) {
                                  results.push({ issueId: item.row.issueId, status: 'Erro', message: e.message });
                                }
                              }
                              const successCount = results.filter((result) => result.status === 'Sucesso').length;
                              const errorCount = results.length - successCount;
                              setImResults(results);
                              await loadImWorksheet();
                              setImWorksheetFeedback({
                                type: errorCount ? 'warning' : 'success',
                                message: errorCount
                                  ? `Atualização concluída com alertas: ${successCount} issue(s) atualizada(s) e ${errorCount} com erro.`
                                  : `Atualização concluída no ACC: ${successCount} issue(s) atualizada(s).`
                              });
                            } catch (updateError) {
                              const message = `Falha ao atualizar issues: ${updateError.message}`;
                              setError(message);
                              setImWorksheetFeedback({ type: 'error', message });
                            } finally {
                              setImLoading(false);
                            }
                          }} disabled={!imSheetPreviewRows.length || imLoading}>Atualizar Issues</button>
                        </div>
                        {imWorksheetFeedback && (
                          <p className={`status ${imWorksheetFeedback.type === 'error' ? 'error-message' : ''}`}>
                            {imWorksheetFeedback.message}
                          </p>
                        )}
                        {!!imSheetPreviewRows.length && <div className="documents-table-wrap eap-table-wrap"><table className="documents-table eap-preview-table"><thead><tr><th>Issue ID</th><th>Título</th><th>Campo</th><th>Valor atual</th><th>Novo valor</th><th>Status</th></tr></thead><tbody>{imSheetPreviewRows.slice(0, 500).map((row, idx) => <tr key={idx}><td>{row.issueId}</td><td>{row.title}</td><td>{row.field}</td><td>{row.oldValue}</td><td>{row.newValue}</td><td>{row.validation}</td></tr>)}</tbody></table></div>}
                      </>
                    )}
                  </>
                )}
                {imMode && imMode !== 'worksheet' && (
                  <>
                    <div className="eap-intro">
                      <div>
                        <p className="eyebrow">Fluxo guiado</p>
                        <h3>{imMode === 'create' ? 'Configurar Issues no ACC' : 'Atualizar Issues no ACC'}</h3>
                        <span>{imMode === 'create' ? 'Use a planilha modelo com a aba ‘Config. issue’ para configurar categorias, tipos e campos personalizados.' : 'Importe a planilha Excel e valide somente a aba “ACS Build”. A aba “EAP” não é usada neste fluxo.'}</span>
                      </div>
                      <label className="eap-upload-button">
                        Carregar planilha
                        <input type="file" accept=".xlsx" onChange={handleImFileChange} />
                      </label>
                    </div>
                    {imMode === 'create' && (
                      <div className="eap-flow-steps" aria-label="Etapas da configuracao de issues">
                        <span className={imFileName ? 'done' : ''}>1. Upload</span>
                        <span className={imPreview ? 'done' : ''}>2. Leitura da aba Config. issue</span>
                        <span className={imPreview ? 'done' : ''}>3. Validar configuração</span>
                        <span className={imResults.length ? 'done' : ''}>4. Aplicar no ACC</span>
                      </div>
                    )}
                    {imMode === 'update' && (
                      <>
                        <label><input type="checkbox" checked={imAllowClearEmpty} onChange={(e)=>setImAllowClearEmpty(e.target.checked)} /> Permitir limpar campos vazios (somente quando explicitamente habilitado)</label>
                        <p className="status">Regra padrão: células vazias não apagam dados no ACC.</p>
                      </>
                    )}
                    {imLoading && <p>Validando planilha...</p>}
                    {imFileName && <p>Arquivo: <strong>{imFileName}</strong></p>}
                    {imPreview && (
                      <>
                        <p>{imMode === 'create' ? (imPreview.message || "Pré-visualização completa. Revise antes de aplicar no ACC.") : `Validação concluída. Linhas: ${imPreview.rows?.length || 0}`}</p>
                        {imMode === 'create' && (
                          <>
                            <p>Nenhuma ação será feita até confirmar em <strong>Aplicar configuração no ACC</strong>.</p>
                            <div className="eap-metrics">
                              <div><span>Categorias</span><strong>{imPreview.categories?.length || 0}</strong></div>
                              <div><span>Tipos</span><strong>{imPreview.types?.length || 0}</strong></div>
                              <div><span>Campos</span><strong>{imPreview.fields?.length || 0}</strong></div>
                              <div><span>Matriz</span><strong>{imPreview.matrix?.length || 0}</strong></div>
                            </div>
                            {!!imPreview.issues?.length && (
                              <div className="documents-table-wrap eap-table-wrap">
                                <table className="documents-table eap-preview-table">
                                  <thead><tr><th>Linha</th><th>Seção</th><th>Status</th><th>Ação prevista</th><th>Motivo</th></tr></thead>
                                  <tbody>{imPreview.issues.map((row, i) => <tr key={`iss-${i}`}><td>{row.line}</td><td>{row.section}</td><td>{row.status}</td><td>{row.action || '-'}</td><td>{row.reason || '-'}</td></tr>)}</tbody>
                                </table>
                              </div>
                            )}
                            <div className="documents-table-wrap eap-table-wrap">
                              <table className="documents-table eap-preview-table">
                                <thead><tr><th>Status</th><th>Categoria ACC</th><th>Ação prevista</th><th>Observação</th></tr></thead>
                                <tbody>{(imPreview.categories || []).map((row, i) => <tr key={`cat-${i}`}><td>{row.status}</td><td>{row.category || '-'}</td><td>{row.action}</td><td>{row.observation || '-'}</td></tr>)}</tbody>
                              </table>
                            </div>
                            <div className="documents-table-wrap eap-table-wrap">
                              <table className="documents-table eap-preview-table">
                                <thead><tr><th>Status</th><th>Código</th><th>Categoria</th><th>Tipo</th><th>Ação</th><th>Observação</th></tr></thead>
                                <tbody>{(imPreview.types || []).map((row, i) => <tr key={`typ-${i}`}><td>{row.status}</td><td>{row.code || '-'}</td><td>{row.category || '-'}</td><td>{row.issueType || '-'}</td><td>{row.action}</td><td>{row.observation || '-'}</td></tr>)}</tbody>
                              </table>
                            </div>
                            <div className="documents-table-wrap eap-table-wrap">
                              <table className="documents-table eap-preview-table">
                                <thead><tr><th>Status</th><th>Campo</th><th>Classificação</th><th>Tipo ACC sugerido</th><th>Obrigatoriedade</th><th>Opções</th><th>Ação</th><th>Observação</th></tr></thead>
                                <tbody>{(imPreview.fields || []).map((row, i) => <tr key={`fld-${i}`}><td>{row.status}</td><td>{row.name || '-'}</td><td>{row.classification || '-'}</td><td>{row.type || '-'}</td><td>{row.required || '-'}</td><td>{(row.options || []).join(', ') || '-'}</td><td>{row.action || '-'}</td><td>{row.observation || '-'}</td></tr>)}</tbody>
                              </table>
                            </div>
                          </>
                        )}
                        {!!imPreview.rows?.length && (
                          <div className="documents-table-wrap eap-table-wrap">
                            <table className="documents-table eap-preview-table">
                              <thead><tr><th>Linha</th><th>Issue ID</th><th>Título</th><th>Status da validação</th><th>Valor atual (ACC)</th><th>Valor novo (planilha)</th><th>Mensagem</th></tr></thead>
                              <tbody>{imPreview.rows.slice(0, 120).map((row) => <tr key={row.id}><td>{row.line}</td><td>{row.issueId || '-'}</td><td>{row.title || '-'}</td><td>{row.validation}</td><td>{row.current ? `Status: ${row.current.status || '-'} | Venc.: ${row.current.dueDate || '-'} | Resp.: ${row.current.assignee || '-'}` : '-'}</td><td>{row.proposed ? `Status: ${row.proposed.status || '-'} | Venc.: ${row.proposed.dueDate || '-'} | Resp.: ${row.proposed.assignee || '-'}` : '-'}</td><td>{(row.errors || []).concat(row.warnings || []).join(' | ')}</td></tr>)}</tbody>
                            </table>
                          </div>
                        )}
                        {!imPreview.rows?.length && imMode !== 'create' && (
                          <div className="empty-state">
                            <strong>Nenhuma linha elegível para prévia.</strong>
                            <span>Verifique a aba/cabeçalhos da planilha e os campos "acao", "titulo" e "issue_id".</span>
                          </div>
                        )}
                        <button className="primary-button" type="button" onClick={executeImFlow} disabled={imLoading || (imMode === 'create' && ((imPreview?.summary?.errors || 0) > 0)) || (imMode === 'update' && !(imPreview?.readyRows > 0))}>
                          {imMode === 'create' ? 'Aplicar configuração no ACC' : 'Atualizar Issues no ACC'}
                        </button>
                        {imActionFeedback && (
                          <p className={`status ${imActionFeedback.type === 'error' ? 'error-message' : ''}`}>
                            {imActionFeedback.message}
                          </p>
                        )}
                        {imMode === 'create' && (
                          <button className="ghost-button" type="button" onClick={downloadImReport}>
                            Baixar relatório da configuração
                          </button>
                        )}
                      </>
                    )}
                    {imResults.length > 0 && (
                      <>
                        {imApplySummary && (
                          <p>
                            <strong>{imApplySummary.persistenceVerified ? 'Persistência confirmada no ACC.' : 'Persistência não confirmada para todos os itens.'}</strong>{' '}
                            {imApplySummary.message}
                          </p>
                        )}
                        <p>Processamento concluído. Registros: {imResults.length}</p>
                        <p>
                          Categorias: criadas {imResults.filter((r) => r.entity === 'categoria' && r.status === 'criado').length}, já existentes {imResults.filter((r) => r.entity === 'categoria' && r.status === 'ja-existe').length}.{' '}
                          Tipos: criados {imResults.filter((r) => r.entity === 'tipo' && r.status === 'criado').length}, já existentes {imResults.filter((r) => r.entity === 'tipo' && r.status === 'ja-existe').length}.{' '}
                          Campos personalizados: criados {imResults.filter((r) => r.entity === 'campo' && r.status === 'criado').length}, já existentes {imResults.filter((r) => r.entity === 'campo' && r.status === 'ja-existe').length}.
                        </p>
                        <div className="documents-table-wrap eap-table-wrap">
                          <table className="documents-table eap-preview-table">
                            <thead><tr><th>Linha</th><th>Entidade</th><th>Nome</th><th>Status</th><th>ID ACC</th><th>Mensagem</th></tr></thead>
                            <tbody>{imResults.map((row, i) => <tr key={`res-${i}`}><td>{row.line || '-'}</td><td>{row.entity || '-'}</td><td>{row.name || '-'}</td><td>{row.status || '-'}</td><td>{row.accId || '-'}</td><td>{row.message || '-'}</td></tr>)}</tbody>
                          </table>
                        </div>
                      </>
                    )}
                  </>
                )}
              </section>
            )}

            {activeModule === 'eap-import' && (
              <section className="eap-panel">
                <div className="eap-intro">
                  <div>
                    <p className="eyebrow">Fluxo seguro</p>
                    <h3>IMPORTACAO DE EAP PARA ISSUES</h3>
                    <span>
                      Primeiro carregue a planilha. O app valida a aba EAP e mostra uma previa antes de qualquer envio ao ACC.
                    </span>
                  </div>
                  <label className="eap-upload-button">
                    Selecionar Excel
                    <input type="file" accept=".xlsx" onChange={handleEapFileChange} />
                  </label>
                </div>

                <div className="eap-flow-steps" aria-label="Etapas da importacao">
                  <span className={eapFileName ? 'done' : ''}>1. Upload</span>
                  <span className={eapPreview ? 'done' : ''}>2. Leitura da aba ACS Build</span>
                  <span className={eapPreview ? 'done' : ''}>3. Validacao</span>
                  <span className={eapResults.length ? 'done' : ''}>4. Resultado</span>
                </div>

                {eapFileName && (
                  <div className="eap-file-card">
                    <div>
                      <p className="eyebrow">Arquivo selecionado</p>
                      <strong>{eapFileName}</strong>
                      {eapPreview && (
                        <span>
                          Aba: {eapPreview.sheetName} (esperado: ACS Build) | Linha de cabecalho: {Number(eapPreview.headerRowIndex || 0) + 1}
                          {Number(eapPreview.friendlyHeaderRowIndex ?? -1) >= 0
                            ? ` | Linha amigavel: ${Number(eapPreview.friendlyHeaderRowIndex) + 1}`
                            : ''}
                        </span>
                      )}
                    </div>
                    <button className="ghost-button" type="button" onClick={resetEapImport}>
                      Limpar
                    </button>
                  </div>
                )}

                {eapLoading && (
                  <div className="loading-box">
                    <strong>Processando planilha ACS Build...</strong>
                    <span>Nenhum issue sera criado durante esta etapa.</span>
                  </div>
                )}

                {eapPreview && (
                  <>
                    <div className="eap-metrics">
                      <div>
                        <span>Total lido</span>
                        <strong>{eapMetrics.total}</strong>
                      </div>
                      <div>
                        <span>Prontos</span>
                        <strong>{eapMetrics.ready}</strong>
                      </div>
                      <div>
                        <span>Alertas</span>
                        <strong>{eapMetrics.incomplete}</strong>
                      </div>
                      <div>
                        <span>Duplicados</span>
                        <strong>{eapMetrics.duplicated}</strong>
                      </div>
                      <div className={eapMetrics.errors ? 'danger-metric' : ''}>
                        <span>Erros</span>
                        <strong>{eapMetrics.errors}</strong>
                      </div>
                      <div>
                        <span>Editadas</span>
                        <strong>{eapMetrics.edited}</strong>
                      </div>
                    </div>

                    <div className="eap-mapping-card">
                      <div>
                        <p className="eyebrow">Colunas identificadas</p>
                        <h3>Mapeamento preliminar</h3>
                        <span>
                          Depois que voce enviar a planilha definitiva, ajustamos os nomes das colunas e os campos personalizados finos.
                        </span>
                      </div>
                      <div className="eap-column-list">
                        {(eapPreview.columns || []).map((column) => (
                          <span key={column.key}>{column.label}</span>
                        ))}
                      </div>
                    </div>

                    <div className="eap-actions">
                      <label className="toggle-row">
                        <input type="checkbox" checked={eapDryRun} onChange={(event) => setEapDryRun(event.target.checked)} />
                        Modo teste: simular sem criar no ACC
                      </label>
                      <button
                        className="primary-button"
                        type="button"
                        disabled={eapLoading || eapMetrics.creatable === 0}
                        onClick={() => {
                          if (!eapDryRun && !window.confirm('Criar issues validas no ACC agora?')) return;
                          createEapIssues();
                        }}
                      >
                        {eapDryRun ? 'Simular criacao' : 'Criar Issues no ACC'}
                      </button>
                      <button className="ghost-button" type="button" onClick={() => setEapEditedRows((rows) => rows.map(validateEapRow))}>
                        Revalidar prévia
                      </button>
                    </div>

                    <div className="documents-table-wrap eap-table-wrap">
                      {eapPreview.rows.length > eapPreviewLimit && (
                        <div className="loading-box compact-note">
                          <strong>Mostrando as primeiras {eapPreviewLimit} linhas na tela.</strong>
                          <span>A validacao considerou {eapPreview.rows.length} linhas. A exibicao foi limitada para manter o app rapido.</span>
                        </div>
                      )}
                      <table className="documents-table eap-preview-table">
                        <thead>
                          <tr>
                            <th>Linha Excel</th>
                            {(eapPreview.columns || []).map((column) => (
                              <th key={column.key} title={column.label}>{column.label}</th>
                            ))}
                            <th>Validacao</th>
                          </tr>
                        </thead>
                        <tbody>
                          {eapVisibleRows.map((row) => (
                            <tr key={row.id} className={`eap-row-${row.validation}`}>
                              <td>{row.line}</td>
                              {(eapPreview.columns || []).map((column) => {
                                const options = eapPreview.dropdownOptionsByColumn?.[column.label] || [];
                                const value = row.sourceValues?.[column.key] ?? '';
                                return (
                                  <td key={`${row.id}-${column.key}`}>
                                    {options.length > 1 ? (
                                      <select value={value} onChange={(event) => updateEapSourceCell(row.id, column.key, event.target.value)}>
                                        <option value="">--</option>
                                        {options.map((option) => <option key={option} value={option}>{option}</option>)}
                                      </select>
                                    ) : (
                                      <input value={value} onChange={(event) => updateEapSourceCell(row.id, column.key, event.target.value)} />
                                    )}
                                  </td>
                                );
                              })}
                              <td>
                                <span className={`validation-pill ${row.validation}`}>{row.validation}</span>
                                {Array.isArray(row.customFieldsFilled) && row.customFieldsFilled.length > 0 && (
                                  <small>{row.customFieldsFilled.length} campos personalizados reconhecidos</small>
                                )}
                                {[...(row.errors || []), ...(row.warnings || [])].length > 0 && (
                                  <small>{[...(row.errors || []), ...(row.warnings || [])].join(' | ')}</small>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </>
                )}

                {eapResults.length > 0 && (
                  <section className="eap-results">
                    <div className="documents-toolbar compact-toolbar">
                      <div>
                        <p className="eyebrow">Retorno</p>
                        <h3>RESULTADO DA IMPORTACAO</h3>
                      </div>
                      <button className="ghost-button" type="button" onClick={exportEapResultsXlsx}>
                        Exportar Excel de retorno
                      </button>
                    </div>
                    {eapResultSummary && (
                      <div className="loading-box compact-note">
                        <strong>
                          {eapDryRun ? 'Simulacao concluida.' : 'Importacao concluida.'} Processadas {eapResultSummary.total} linhas.
                        </strong>
                        <span>
                          Criadas: {eapResultSummary.created} | Duplicadas: {eapResultSummary.duplicated} | Erros: {eapResultSummary.errors}
                          {eapResultSummary.simulated ? ` | Simuladas: ${eapResultSummary.simulated}` : ''}
                          {eapResultSummary.review ? ` | Revisar: ${eapResultSummary.review}` : ''}
                        </span>
                      </div>
                    )}
                    <div className="documents-table-wrap">
                      <table className="documents-table eap-result-table">
                        <thead>
                          <tr>
                            <th>Linha</th>
                            <th>Codigo EAP</th>
                            <th>Nome do Issue</th>
                            <th>Status da criacao</th>
                            <th>ID ACC</th>
                            <th>Link</th>
                            <th>Mensagem</th>
                          </tr>
                        </thead>
                        <tbody>
                          {eapResults.map((row) => (
                            <tr key={`${row.line}-${row.title}`}>
                              <td>{row.line}</td>
                              <td>{row.eapCode || '-'}</td>
                              <td>{row.title || '-'}</td>
                              <td>{row.status || row.validation}</td>
                              <td>{row.issueId || '-'}</td>
                              <td>
                                {row.issueLink ? (
                                  <a href={row.issueLink} target="_blank" rel="noreferrer">
                                    Abrir
                                  </a>
                                ) : (
                                  '-'
                                )}
                              </td>
                              <td>{row.message || row.error || '-'}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                    {eapExecutionHistory.length > 0 && (
                      <div className="eap-log-card">
                        <p className="eyebrow">Registro rapido</p>
                        <h4>Historico de execucoes</h4>
                        <ul>
                          {eapExecutionHistory.map((entry) => (
                            <li key={entry.id}>
                              {formatDateTime(entry.createdAt)} - {entry.dryRun ? 'Simulacao' : 'Criacao real'}: {entry.created} criadas de{' '}
                              {entry.total} linhas (duplicadas: {entry.duplicated}, erros: {entry.errors}).
                            </li>
                          ))}
                        </ul>
                      </div>
                    )}
                  </section>
                )}

                {!eapPreview && !eapLoading && (
                  <div className="empty-state">
                    <strong>Nenhuma planilha ACS Build carregada ainda.</strong>
                    <span>Escolha um arquivo .xlsx para ver a previa. O upload sozinho nao cria issues.</span>
                  </div>
                )}
              </section>
            )}

            {activeModule === 'impact' && (
              <section className="impact-panel">
                <div className="impact-toolbar">
                  <div>
                    <p className="eyebrow">Campo analisado</p>
                    <h3>{selectedImpactFieldLabel}</h3>
                    <span>Issues organizadas por prazo e criticidade do campo personalizado selecionado.</span>
                  </div>
                  <label>
                    Selecionar impacto
                    <select value={selectedImpactField} onChange={(event) => setSelectedImpactField(event.target.value)}>
                      {impactFields.map((field) => (
                        <option key={field.value} value={field.value}>
                          {field.label}
                        </option>
                      ))}
                    </select>
                  </label>
                  <div className="impact-filters">
                    <label>
                      Data inicial
                      <input type="date" value={impactStartDate} onChange={(event) => setImpactStartDate(event.target.value)} />
                    </label>
                    <label>
                      Data final
                      <input type="date" value={impactEndDate} onChange={(event) => setImpactEndDate(event.target.value)} />
                    </label>
                    <label>
                      Categoria *
                      <select value={impactCategoryFilter} onChange={(event) => setImpactCategoryFilter(event.target.value)}>
                        <option value="all">Todas as categorias</option>
                        {impactCategoryOptions.map((category) => (
                          <option key={category} value={category}>
                            {category}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      Tipo de issue
                      <select value={impactTypeFilter} onChange={(event) => setImpactTypeFilter(event.target.value)}>
                        <option value="all">Todos os tipos</option>
                        {impactTypeOptions.map((type) => (
                          <option key={type} value={type}>
                            {type}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label>
                      Fase
                      <select value={impactPhaseFilter} onChange={(event) => setImpactPhaseFilter(event.target.value)}>
                        <option value="all">Todas as fases</option>
                        {impactPhaseOptions.map((phase) => (
                          <option key={phase} value={phase}>
                            {phase}
                          </option>
                        ))}
                      </select>
                    </label>
                    <button
                      className="ghost-button"
                      type="button"
                      onClick={() => {
                        setImpactStartDate('');
                        setImpactEndDate('');
                        setImpactCategoryFilter('all');
                        setImpactTypeFilter('all');
                        setImpactPhaseFilter('all');
                      }}
                    >
                      Limpar filtros
                    </button>
                  </div>
                </div>

                <div className="impact-metrics">
                  <div className="impact-metric-total">
                    <i className="impact-metric-icon" aria-hidden="true" />
                    <span>Total analisado</span>
                    <strong>{impactMetrics.total}</strong>
                  </div>
                  {impactMetrics.byLevel.map((level) => (
                    <div key={level.value} className={`impact-metric-${level.key}`} style={{ '--impact-color': level.color }}>
                      <i className="impact-metric-icon" aria-hidden="true" />
                      <span>{level.label}</span>
                      <strong>{level.total}</strong>
                    </div>
                  ))}
                  <div className="impact-metric-date">
                    <i className="impact-metric-icon" aria-hidden="true" />
                    <span>Sem prazo</span>
                    <strong>{impactMetrics.withoutDueDate}</strong>
                  </div>
                </div>

                <div className="impact-layout">
                  <div className="impact-chart">
                    <div className="impact-chart-title">
                      <div>
                        <h3>Linha de impactos</h3>
                        <span>Evolução dos issues por prazo e criticidade</span>
                      </div>
                    </div>
                    <div className="impact-timeline">
                      <div className="impact-axis-labels" aria-hidden="true">
                        {impactTimelineData.ticks.map((tick) => (
                          <span key={tick.key} style={{ left: `${tick.position}%` }}>
                            {tick.label}
                          </span>
                        ))}
                      </div>
                      {impactTimelineData.ticks.map((tick) => (
                        <span key={tick.key} className="impact-axis-line" style={{ left: `${tick.position}%` }} aria-hidden="true" />
                      ))}
                      {impactTimelineData.lanes.map((level) => (
                        <div key={level.value} className={`impact-timeline-row impact-timeline-row-${level.key}`} style={{ '--impact-color': level.color }}>
                          <div className="impact-row-label">
                            <i aria-hidden="true" />
                            <span>{level.label}</span>
                          </div>
                          <div className="impact-row-track">
                            <span className="impact-row-line" aria-hidden="true" />
                            {level.issues.map((issue, issueIndex) => (
                              <button
                                key={issue.id}
                                type="button"
                                className={`impact-bubble impact-bubble-${level.key} ${selectedIssueId === issue.id ? 'selected' : ''}`}
                                style={{ left: `${issue.position}%`, '--impact-color': level.color }}
                                onClick={() => setSelectedIssueId(issue.id)}
                                title={`${issue.title} - ${formatDate(issue.dueDate)}`}
                              >
                                {issue.displayId || issueIndex + 1}
                              </button>
                            ))}
                          </div>
                        </div>
                      ))}
                      {impactTimelineData.withoutDueDate.length > 0 && (
                        <div className="impact-no-date-row">
                          <span>Sem prazo definido</span>
                          <div>
                            {impactTimelineData.withoutDueDate.map((issue) => (
                              <button
                                key={issue.id}
                                type="button"
                                className={`impact-no-date-chip ${selectedIssueId === issue.id ? 'selected' : ''}`}
                                onClick={() => setSelectedIssueId(issue.id)}
                              >
                                {issue.displayId || issue.title}
                              </button>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                    <div className="impact-legend">
                      {impactTimelineData.lanes.map((level) => (
                        <span key={level.value}><i style={{ '--impact-color': level.color }} />{level.label}</span>
                      ))}
                      <span><i className="impact-dot-selected" />Selecionado</span>
                    </div>
                  </div>

                  <aside className="impact-detail-card">
                    {selectedIssueFull ? (
                      <>
                        <p className="eyebrow">Issue selecionada</p>
                        <h3>{selectedIssueFull.title}</h3>
                        <div className="impact-detail-grid">
                          <span><em>ID:</em> <strong>{selectedIssueFull.displayId || selectedIssueFull.id}</strong></span>
                          <span><em>Status:</em> <strong>{selectedIssueFull.status}</strong></span>
                          <span><em>Responsável:</em> <strong>{formatDetailValue(selectedIssueFull.assignedTo)}</strong></span>
                          <span><em>Criada em:</em> <strong>{formatDateTime(selectedIssueFull.createdAt)}</strong></span>
                          <span><em>Prazo:</em> <strong>{selectedIssueFull.dueDate ? formatDate(selectedIssueFull.dueDate) : 'Sem prazo definido'}</strong></span>
                          <span><em>Tipo:</em> <strong>{selectedIssueFull.issueType || 'Sem tipo'}</strong></span>
                          <span><em>Subtipo:</em> <strong>{selectedIssueFull.issueSubtype || 'Sem subtipo'}</strong></span>
                          <span><em>Campo analisado:</em> <strong>{selectedImpactFieldLabel}</strong></span>
                          <span><em>{usesImpactOptionValues ? 'Valor identificado:' : 'Classificação:'}</em> <strong>{usesImpactOptionValues ? normalizeImpactOption(getImpactFieldValue(selectedIssueFull, selectedImpactField)) : normalizeImpactLevel(getImpactFieldValue(selectedIssueFull, selectedImpactField))}</strong></span>
                          <span><em>Localização:</em> <strong>{formatDetailValue(selectedIssueFull.location)}</strong></span>
                        </div>
                        {selectedIssueFull.description && (
                          <div className="issue-description">
                            <p className="label">Descrição</p>
                            <p>{selectedIssueFull.description}</p>
                          </div>
                        )}
                        {selectedIssueFull.raw?.links?.webView?.href && (
                          <a className="ghost-button impact-acc-link" href={selectedIssueFull.raw.links.webView.href} target="_blank" rel="noreferrer">
                            Abrir no ACC
                          </a>
                        )}
                        {selectedIssueDetailsLoading && <p className="muted">Carregando ficha completa...</p>}
                      </>
                    ) : (
                      <div className="empty-state compact-empty">
                        <strong>Selecione uma issue no gráfico.</strong>
                        <span>Ao clicar em uma bolha, os detalhes aparecerão aqui.</span>
                      </div>
                    )}
                  </aside>
                </div>
              </section>
            )}

            {activeModule === 'bim-awp' && (
              <section className="bim-awp-panel">
                <div className="bim-awp-hero">
                  <span className="bim-awp-hero-icon" aria-hidden="true">
                    <UiIcon name="bim" />
                  </span>
                  <div>
                    <p className="eyebrow">Modelo BIM + AWP</p>
                    <h3>Coordenação BIM</h3>
                    <p>
                      Painel para acompanhar modelos, workfaces, validação, qualidade,
                      interferências e status de publicação usando os dados já lidos do ACC.
                    </p>
                    <ul>
                      <li>Modelo BIM e workfaces</li>
                      <li>Validação, qualidade e interferências</li>
                      <li>Status de publicação e documentos emitidos</li>
                    </ul>
                  </div>
                  <div className="bim-awp-score">
                    <strong>{bimAwpDashboard.total}</strong>
                    <span>issues relacionados</span>
                  </div>
                </div>

                <div className="bim-awp-metrics">
                  <div>
                    <span>Em aberto</span>
                    <strong>{bimAwpDashboard.open}</strong>
                  </div>
                  <div className={bimAwpDashboard.overdue > 0 ? 'danger-metric' : ''}>
                    <span>Atrasados</span>
                    <strong>{bimAwpDashboard.overdue}</strong>
                  </div>
                  <div>
                    <span>Workfaces</span>
                    <strong>{bimAwpDashboard.workfaces}</strong>
                  </div>
                  <div>
                    <span>Modelo BIM / AWP</span>
                    <strong>{bimAwpDashboard.modelIssues}</strong>
                  </div>
                  <div>
                    <span>Validação / qualidade</span>
                    <strong>{bimAwpDashboard.validationIssues}</strong>
                  </div>
                  <div>
                    <span>Publicação</span>
                    <strong>{bimAwpDashboard.publicationIssues}</strong>
                  </div>
                </div>

                <div className="bim-awp-grid">
                  <article className="dashboard-card">
                    <h4>Status dos issues BIM/AWP</h4>
                    {bimAwpDashboard.statusGroups.length ? (
                      <div className="bim-awp-status-list">
                        {bimAwpDashboard.statusGroups.map((status) => (
                          <span key={status.name}>
                            <em>{status.name}</em>
                            <strong>{status.value}</strong>
                          </span>
                        ))}
                      </div>
                    ) : (
                      <p className="muted">Nenhum issue BIM/AWP identificado no filtro atual.</p>
                    )}
                  </article>
                  <article className="dashboard-card">
                    <h4>Status de publicação</h4>
                    <div className="bim-awp-publication">
                      <span>{bimAwpDashboard.documentsEmitted} de {bimAwpDashboard.documentsTotal} documentos emitidos</span>
                      <div aria-hidden="true">
                        <i style={{ width: `${bimAwpDashboard.documentsPercent}%` }} />
                      </div>
                      <small>{bimAwpDashboard.documentsPending} pendentes na lista mestre</small>
                    </div>
                  </article>
                </div>

                <article className="bim-parameter-card">
                  <div className="bim-parameter-heading">
                    <div>
                      <p className="eyebrow">Parâmetros de modelo</p>
                      <h3>Plano para RVT e DWG</h3>
                      <span>
                        A Autodesk oferece Extended Properties no Model Coordination. Esta tela organiza os parâmetros para validação
                        e exportação enquanto conectamos a escrita direta via API quando o projeto tiver a função habilitada.
                      </span>
                    </div>
                    <button className="ghost-button" type="button" onClick={exportBimParameterPlanCsv}>
                      Exportar plano CSV
                    </button>
                  </div>

                  <form className="bim-parameter-form" onSubmit={addBimParameterPlanItem}>
                    <label>
                      Nome do parâmetro
                      <input
                        type="text"
                        value={bimParameterDraft.name}
                        onChange={(event) => handleBimParameterDraftChange('name', event.target.value)}
                        placeholder="Ex.: Código EAP, Workface, Status de publicação"
                      />
                    </label>
                    <label>
                      Tipo
                      <select
                        value={bimParameterDraft.type}
                        onChange={(event) => handleBimParameterDraftChange('type', event.target.value)}
                      >
                        <option>Texto</option>
                        <option>Número</option>
                        <option>Data</option>
                        <option>Lista suspensa</option>
                        <option>Sim/Não</option>
                      </select>
                    </label>
                    <label>
                      Aplicação
                      <select
                        value={bimParameterDraft.target}
                        onChange={(event) => handleBimParameterDraftChange('target', event.target.value)}
                      >
                        <option>RVT e DWG</option>
                        <option>Somente RVT</option>
                        <option>Somente DWG</option>
                        <option>Elementos coordenados</option>
                      </select>
                    </label>
                    <label>
                      Valor padrão
                      <input
                        type="text"
                        value={bimParameterDraft.value}
                        onChange={(event) => handleBimParameterDraftChange('value', event.target.value)}
                        placeholder="Opcional"
                      />
                    </label>
                    <label className="bim-parameter-notes">
                      Observações
                      <input
                        type="text"
                        value={bimParameterDraft.notes}
                        onChange={(event) => handleBimParameterDraftChange('notes', event.target.value)}
                        placeholder="Ex.: preencher por disciplina, workface ou pacote"
                      />
                    </label>
                    <button className="primary-button" type="submit">
                      Adicionar ao plano
                    </button>
                  </form>
                  {bimParameterMessage && <p className="bim-parameter-message">{bimParameterMessage}</p>}

                  {bimParameterPlan.length > 0 && (
                    <div className="bim-parameter-table-wrap">
                      <table className="documents-table bim-parameter-table">
                        <thead>
                          <tr>
                            <th>Parâmetro</th>
                            <th>Tipo</th>
                            <th>Aplicação</th>
                            <th>Valor padrão</th>
                            <th>Observações</th>
                            <th>Status</th>
                          </tr>
                        </thead>
                        <tbody>
                          {bimParameterPlan.map((item) => (
                            <tr key={item.id}>
                              <td>{item.name}</td>
                              <td>{item.type}</td>
                              <td>{item.target}</td>
                              <td>{item.value || '-'}</td>
                              <td>{item.notes || '-'}</td>
                              <td><span className="status-pill status-open">{item.status}</span></td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </article>
              </section>
            )}

            {activeModule === 'project-management' && (
              <section className="project-management-panel">
                <article className="permission-note">
                  <span className="module-icon" aria-hidden="true"><UiIcon name="project" /></span>
                  <div>
                    <p className="eyebrow">Segurança por usuário</p>
                    <h3>Marcos contratuais</h3>
                    <p>
                      Esta visão usa somente as issues que o ACC retorna para a conta Autodesk logada.
                      Se uma issue não aparece para o usuário no ACC, ela também não entra neste quadro.
                    </p>
                  </div>
                </article>

                <div className="metric-grid project-management-metrics">
                  <div>
                    <span>Marcos visíveis</span>
                    <strong>{projectManagementMetrics.milestones}</strong>
                  </div>
                  <div>
                    <span>Issues visíveis</span>
                    <strong>{projectManagementMetrics.visibleIssues}</strong>
                  </div>
                  <div>
                    <span>Medidos</span>
                    <strong>{projectManagementMetrics.medidos}</strong>
                  </div>
                  <div>
                    <span>Execução geral</span>
                    <strong>{projectManagementMetrics.execution}%</strong>
                  </div>
                  <div>
                    <span>A medir</span>
                    <strong>{projectManagementMetrics.aMedir}</strong>
                  </div>
                  <div className={projectManagementMetrics.highRisk > 0 ? 'danger-metric' : ''}>
                    <span>Risco alto</span>
                    <strong>{projectManagementMetrics.highRisk}</strong>
                  </div>
                </div>

                <div className="project-management-filters">
                  <label>
                    Pesquisar
                    <input
                      type="search"
                      value={projectManagementSearch}
                      onChange={(event) => setProjectManagementSearch(event.target.value)}
                      placeholder="Buscar por marco, entregável, valor ou status..."
                    />
                  </label>
                  <label>
                    Status medição
                    <select value={projectManagementStatusFilter} onChange={(event) => setProjectManagementStatusFilter(event.target.value)}>
                      <option value="all">Todos</option>
                      {projectManagementStatusOptions.map((status) => (
                        <option key={status} value={status}>{status}</option>
                      ))}
                    </select>
                  </label>
                  <button
                    className="ghost-button"
                    type="button"
                    onClick={() => {
                      setProjectManagementSearch('');
                      setProjectManagementStatusFilter('all');
                    }}
                  >
                    Limpar filtros
                  </button>
                </div>

                {projectManagementRows.length > 0 ? (
                  <div className="documents-table-wrap project-management-table-wrap">
                    <table className="documents-table project-management-table">
                      <thead>
                        <tr>
                          <th>Marco</th>
                          <th>Entregáveis</th>
                          <th>Valor da medição</th>
                          <th>Prazo contratual</th>
                          <th>Data prevista G5</th>
                          <th>Avanço</th>
                          <th>Risco</th>
                          <th>Status medição</th>
                          <th>Issues visíveis</th>
                        </tr>
                      </thead>
                      <tbody>
                        {projectManagementRows.map((row) => (
                          <tr key={row.key} onClick={() => row.issues[0] && setSelectedIssueId(row.issues[0].id)}>
                            <td><strong>{row.marco}</strong></td>
                            <td>
                              <div className="milestone-deliverable-summary">
                                <strong>{row.milestoneTitle}</strong>
                                <span>{row.issueCount} issues vinculadas a este marco</span>
                                <div className="milestone-summary-chips">
                                  <span><strong>{row.statusSummary.concluidas}</strong> concluídas</span>
                                  <span><strong>{row.statusSummary.abertas}</strong> abertas</span>
                                  <span><strong>{row.statusSummary.andamento}</strong> em andamento</span>
                                  <span><strong>{row.statusSummary.pendentes}</strong> pendentes</span>
                                </div>
                              </div>
                            </td>
                            <td>
                              <input
                                className="currency-input"
                                defaultValue={row.valueInput}
                                onClick={(event) => event.stopPropagation()}
                                onBlur={(event) => updateProjectManagementMeasurementValue(row, event.target.value)}
                                disabled={savingIssueId === row.measurementIssueId}
                                aria-label={`Valor da medição do marco ${row.marco}`}
                              />
                            </td>
                            <td>{row.contractualDate ? formatDate(row.contractualDate) : '-'}</td>
                            <td>{row.expectedDate ? formatDate(row.expectedDate) : '-'}</td>
                            <td>
                              <div className="progress-cell" aria-label={`Avanço calculado em ${row.executionPercent}%`}>
                                <span>{row.executionPercent}%</span>
                                <div className="progress-bar"><i style={{ width: `${row.executionPercent}%` }} /></div>
                              </div>
                            </td>
                            <td>
                              <span className={`risk-chip risk-${normalizeText(row.risk).replace(/\s+/g, '-')}`}>
                                {row.risk}
                              </span>
                            </td>
                            <td>
                              <select
                                className="measurement-status-select"
                                value={normalizeText(row.statusMedicao).includes('medido') ? 'Medido' : 'A medir'}
                                onClick={(event) => event.stopPropagation()}
                                onChange={(event) => updateProjectManagementMeasurementStatus(row, event.target.value)}
                                disabled={savingIssueId === row.measurementIssueId}
                                aria-label={`Status de medição do marco ${row.marco}`}
                              >
                                <option value="A medir">A medir</option>
                                <option value="Medido">Medido</option>
                              </select>
                            </td>
                            <td>{row.issueCount}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                ) : (
                  <div className="empty-state">
                    <strong>Nenhum marco contratual visível para este usuário.</strong>
                    <span>
                      Confirme se existem issues com campos de marco, valor ou status de medição e se a pessoa logada tem permissão para vê-las no ACC.
                    </span>
                  </div>
                )}

                <section className="project-management-dashboard">
                  <div className="cronograma-dashboard-header">
                    <div>
                      <h3>Dashboard dos marcos</h3>
                      <p className="muted">Resumo do resultado filtrado por valor, status de medição, risco e volume de issues.</p>
                    </div>
                  </div>
                  <div className="dashboard-grid two-cols">
                    <article className="dashboard-card project-value-card">
                      <h4>Valor da medição</h4>
                      <ul className="pred-list">
                        <li><span>Total visível</span><strong>{formatCurrencyBRL(projectManagementMetrics.totalValue)}</strong></li>
                        <li><span>Medido</span><strong>{formatCurrencyBRL(projectManagementMetrics.measuredValue)}</strong></li>
                        <li><span>A medir</span><strong>{formatCurrencyBRL(projectManagementMetrics.pendingValue)}</strong></li>
                      </ul>
                    </article>
                    <article className="dashboard-card">
                      <h4>Status de medição</h4>
                      {projectManagementMetrics.statusData.map((item) => (
                        <div key={item.name} className="stack-row">
                          <span>{item.name}</span>
                          <div><i style={{ width: `${(item.value / Math.max(1, projectManagementMetrics.milestones)) * 100}%`, background: item.color }} /></div>
                          <strong>{item.value}</strong>
                        </div>
                      ))}
                    </article>
                  </div>
                  <div className="dashboard-grid two-cols">
                    <article className="dashboard-card">
                      <h4>Risco por marco</h4>
                      {projectManagementMetrics.riskData.map((item) => (
                        <div key={item.name} className="stack-row">
                          <span>{item.name}</span>
                          <div><i style={{ width: `${(item.value / Math.max(1, projectManagementMetrics.milestones)) * 100}%`, background: item.color }} /></div>
                          <strong>{item.value}</strong>
                        </div>
                      ))}
                    </article>
                    <article className="dashboard-card">
                      <h4>Issues por marco</h4>
                      {projectManagementMetrics.topMilestones.map((item) => (
                        <div key={item.key} className="bar-row">
                          <span>{item.marco}</span>
                          <div><i style={{ width: `${(item.issueCount / Math.max(1, projectManagementMetrics.visibleIssues)) * 100}%`, background: '#0f7c90' }} /></div>
                          <strong>{item.issueCount}</strong>
                        </div>
                      ))}
                    </article>
                  </div>
                </section>
              </section>
            )}

            {activeModule === 'coordination' && (
              <>
            <div className="metric-grid">
              <div>
                <span>Total no projeto</span>
                <strong>{metrics.total}</strong>
              </div>
              <div>
                <span>Resultado filtrado</span>
                <strong>{metrics.visible}</strong>
              </div>
              <div>
                <span>Em aberto</span>
                <strong>{metrics.open}</strong>
              </div>
              <div className={metrics.overdue > 0 ? 'danger-metric' : ''}>
                <span>Atrasadas</span>
                <strong>{metrics.overdue}</strong>
              </div>
            </div>

            <div className="filter-bar" aria-label="Filtros de issues">
              <button type="button" className={issueFilter === 'all' ? 'active' : ''} onClick={() => setIssueFilter('all')}>
                Todas
              </button>
              <button
                type="button"
                className={issueFilter === 'overdue' ? 'active' : ''}
                onClick={() => setIssueFilter('overdue')}
              >
                Atrasadas
              </button>
            </div>

            <div className="advanced-filters">
              <label>
                Categoria
                <select value={categoryFilter} onChange={(event) => setCategoryFilter(event.target.value)}>
                  <option value="all">Todas</option>
                  {filterOptions.categories.map((category) => (
                    <option key={category} value={category}>
                      {category}
                    </option>
                  ))}
                </select>
              </label>

              <label>
                Tipo de issue
                <select value={typeFilter} onChange={(event) => setTypeFilter(event.target.value)}>
                  <option value="all">Todos</option>
                  {filterOptions.types.map((type) => (
                    <option key={type} value={type}>
                      {type}
                    </option>
                  ))}
                </select>
              </label>

              <button
                type="button"
                className="clear-filter-button"
                onClick={() => {
                  setIssueFilter('all');
                  setCategoryFilter('all');
                  setTypeFilter('all');
                }}
              >
                Limpar filtros
              </button>
            </div>

            {issuesLoading && (
              <div className="loading-box">
                <strong>Buscando issues do projeto selecionado...</strong>
                <span>Se houver erro de permissão da Autodesk, ele aparecerá logo abaixo.</span>
              </div>
            )}
            {!issuesLoading && visibleIssues.length === 0 && (
              <div className="empty-state">
                <strong>Nenhuma issue encontrada para o filtro selecionado.</strong>
                <span>Use o filtro Todas para ver as issues disponíveis, inclusive quando a Autodesk não envia categoria, tipo ou subtipo.</span>
              </div>
            )}

            {monthlyTimeline.length > 0 && (
              <div className="timeline-panel" aria-label="Linha do tempo mensal de issues">
                {monthlyTimeline.map((month) => (
                  <div
                    key={month.monthKey}
                    className={`timeline-month ${month.overdue > 0 ? 'has-overdue' : ''}`}
                  >
                    <span className="timeline-marker" />
                    <div className="timeline-content">
                      <div className="timeline-header">
                        <strong>{month.label}</strong>
                        <span>{month.total} issues</span>
                      </div>
                      <p>
                        {month.open} abertas
                        {month.overdue > 0 ? ` • ${month.overdue} atrasadas` : ' • nenhuma atrasada'}
                      </p>
                      <ul>
                        {month.issues.map((issue) => (
                          <li key={issue.id}>
                            <button
                              type="button"
                              className={isOverdue(issue) ? 'overdue-title' : ''}
                              onClick={() => setSelectedIssueId(issue.id)}
                            >
                              {issue.title}
                            </button>
                          </li>
                        ))}
                      </ul>
                    </div>
                  </div>
                ))}
              </div>
            )}

            {selectedIssue && (
              <section className="issue-detail-panel">
                <div className="issue-detail-title">
                  <p className="eyebrow">Issue selecionada</p>
                  <h3>{selectedIssueFull.title}</h3>
                  <p className="muted">ID Autodesk: {selectedIssueFull.id}</p>
                  {selectedIssueFull.displayId && <p className="muted">Identificador: {selectedIssueFull.displayId}</p>}
                  {selectedIssueDetailsLoading && <p className="muted">Carregando ficha completa...</p>}
                  {selectedIssueDetailsError && (
                    <p className="warning-text">Nao consegui carregar todos os detalhes: {selectedIssueDetailsError}</p>
                  )}
                </div>
                <div className="detail-grid">
                  {selectedIssueRows.map(([label, value]) => (
                    <span key={`${selectedIssueFull.id}-${label}`}>
                      {label}: <strong>{formatDetailValue(value)}</strong>
                    </span>
                  ))}
                </div>
                <div className="issue-controls selected-issue-controls">
                  <label className="wide-field">
                    Título
                    <input
                      defaultValue={selectedIssueFull.title || ''}
                      onBlur={(event) => updateIssue(selectedIssue.id, { title: event.target.value })}
                      disabled={savingIssueId === selectedIssue.id}
                    />
                  </label>

                  <label>
                    Status
                    <select
                      value={selectedIssue?.status || ''}
                      onChange={(event) => updateIssue(selectedIssue.id, { status: event.target.value })}
                      disabled={savingIssueId === selectedIssue.id}
                    >
                      <option value={selectedIssue?.status || ''}>{selectedIssue?.status || 'Selecione'}</option>
                      {statusOptions.map((option) => (
                        <option key={option.value} value={option.value}>
                          {option.label}
                        </option>
                      ))}
                    </select>
                  </label>

                  <label>
                    Prazo
                    <input
                      type="date"
                      defaultValue={getDateInputValue(selectedIssue?.dueDate)}
                      onBlur={(event) => updateIssue(selectedIssue.id, { dueDate: event.target.value })}
                      disabled={savingIssueId === selectedIssue.id}
                    />
                  </label>

                  <label>
                    Responsável
                    <select
                      value={selectedIssueFull.assignedToId || selectedIssueFull.raw?.assignedTo || ''}
                      onChange={(event) => updateIssue(selectedIssue.id, { assignedTo: event.target.value })}
                      disabled={savingIssueId === selectedIssue.id}
                    >
                      <option value="">Sem responsável</option>
                      {projectUsers.map((projectUser) => (
                        <option key={projectUser.id} value={projectUser.id}>
                          {projectUser.name}
                          {projectUser.email ? ` - ${projectUser.email}` : ''}
                        </option>
                      ))}
                    </select>
                  </label>
                </div>
                <div className="issue-description">
                  <label className="label" htmlFor={`description-${selectedIssueFull.id}`}>Descrição</label>
                  <textarea
                    id={`description-${selectedIssueFull.id}`}
                    defaultValue={selectedIssueFull.description || ''}
                    placeholder="Descreva a solicitação, pendência, decisão ou ação esperada."
                    onBlur={(event) => updateIssue(selectedIssue.id, { description: event.target.value })}
                    disabled={savingIssueId === selectedIssue.id}
                  />
                </div>
                {selectedIssueFull.customAttributes?.length > 0 && (
                  <div className="custom-fields selected-custom-fields">
                    <p className="label">Informações:</p>
                    <dl>
                      {selectedIssueFull.customAttributes.map((field) => (
                        <div key={`${selectedIssueFull.id}-${field.id}-${field.name}`}>
                          <dt>{field.name}</dt>
                          <dd>
                            {getCustomFieldOptions(field).length > 0 ? (
                              <select
                                defaultValue={field.rawValue ?? field.value ?? ''}
                                onBlur={(event) => updateCustomAttribute(selectedIssueFull, field, event.target.value)}
                                disabled={savingIssueId === selectedIssue.id}
                              >
                                <option value="">Não informado</option>
                                {getCustomFieldOptions(field).map((option) => (
                                  <option key={option.id} value={option.id}>
                                    {option.label}
                                  </option>
                                ))}
                              </select>
                            ) : (
                              <input
                                defaultValue={field.value ?? ''}
                                onBlur={(event) => updateCustomAttribute(selectedIssueFull, field, event.target.value)}
                                disabled={savingIssueId === selectedIssue.id}
                              />
                            )}
                          </dd>
                        </div>
                      ))}
                    </dl>
                  </div>
                )}
                <div className="issue-comments">
                  <p className="label">Troca de comentarios</p>
                  {selectedIssueFull.commentsWarning && (
                    <p className="warning-text">A Autodesk nao retornou os comentarios: {selectedIssueFull.commentsWarning}</p>
                  )}
                  {selectedIssueFull.comments?.length > 0 ? (
                    <div className="comment-list">
                      {selectedIssueFull.comments.map((comment) => (
                        <article key={comment.id} className="comment-card">
                          <div>
                            <strong>{comment.createdBy || 'Autor nao informado'}</strong>
                            <span>{formatDateTime(comment.createdAt)}</span>
                          </div>
                          <p>{comment.body || 'Comentario sem texto retornado pela Autodesk.'}</p>
                        </article>
                      ))}
                    </div>
                  ) : (
                    <p className="muted">
                      {selectedIssueDetailsLoading
                        ? 'Buscando comentarios...'
                        : 'Nenhum comentario retornado para esta issue.'}
                    </p>
                  )}
                </div>
              </section>
            )}

            {!selectedIssue && (
            <div className="issue-groups" aria-label="Lista de issues agrupadas por tipo">
              {groupedIssues.map((group) => (
                <section key={group.name} className="issue-group">
                  <div className="issue-group-header">
                    <h3>{group.name}</h3>
                    <span>{group.issues.length} issues</span>
                  </div>

                  <div className="issue-table" role="table" aria-label={`Issues de ${group.name}`}>
                    <div className="issue-table-header" role="row">
                      <span>Título</span>
                      <span>Status</span>
                      <span>Categoria</span>
                      <span>Prazo</span>
                      <span>Ação</span>
                    </div>

                    {group.issues.map((issue) => (
                      <article
                        key={issue.id}
                        className={[
                          'issue-row',
                          isOverdue(issue) ? 'overdue' : '',
                          selectedIssueId === issue.id ? 'selected-issue' : ''
                        ].join(' ')}
                        role="row"
                      >
                        <div className="issue-title-cell">
                          <div>
                            <h3>{issue.title}</h3>
                            <p className="muted">ID Autodesk: {issue.id}</p>
                          </div>
                        </div>
                        <span className="status-pill">{issue.status}</span>
                        <span>{issue.category || 'Sem categoria'}</span>
                        <span>{formatDate(issue.dueDate)}</span>
                        <button type="button" className="text-button" onClick={() => setSelectedIssueId(issue.id)}>
                          Ver detalhes
                        </button>

                        <div className="issue-controls">
                          <label>
                            Status
                            <select
                              value={issue.status || ''}
                              onChange={(event) => updateIssue(issue.id, { status: event.target.value })}
                              disabled={savingIssueId === issue.id}
                            >
                              <option value={issue.status || ''}>{issue.status || 'Selecione'}</option>
                              {statusOptions.map((option) => (
                                <option key={option.value} value={option.value}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                          </label>

                          <label>
                            Prazo
                            <input
                              type="date"
                              defaultValue={getDateInputValue(issue.dueDate)}
                              onBlur={(event) => updateIssue(issue.id, { dueDate: event.target.value })}
                              disabled={savingIssueId === issue.id}
                            />
                          </label>

                          <div>
                            <span className="label">Categoria</span>
                            <p>{issue.category || 'Sem categoria'}</p>
                          </div>

                          <div>
                            <span className="label">Tipo de issue</span>
                            <p>{issue.issueType || 'Não informado'}</p>
                          </div>

                          <div>
                            <span className="label">Responsável</span>
                            <p>{issue.assignedTo || 'Não informado'}</p>
                          </div>

                          <div>
                            <span className="label">Vencimento</span>
                            <p>{formatDate(issue.dueDate)}</p>
                          </div>
                        </div>

                        {issue.customAttributes?.length > 0 && (
                          <div className="custom-fields">
                            <p className="label">Informações:</p>
                            <dl>
                              {issue.customAttributes.map((field) => (
                                <div key={`${issue.id}-${field.id}-${field.name}`}>
                                  <dt>{field.name}</dt>
                                  <dd>
                                    {field.options?.length > 0 ? (
                                      <select
                                        defaultValue={field.rawValue ?? ''}
                                        onBlur={(event) => updateCustomAttribute(issue, field, event.target.value)}
                                        disabled={savingIssueId === issue.id}
                                      >
                                        <option value="">Não informado</option>
                                        {field.options.map((option) => (
                                          <option key={option.id} value={option.id}>
                                            {option.label}
                                          </option>
                                        ))}
                                      </select>
                                    ) : (
                                      <input
                                        defaultValue={field.value ?? ''}
                                        onBlur={(event) => updateCustomAttribute(issue, field, event.target.value)}
                                        disabled={savingIssueId === issue.id}
                                      />
                                    )}
                                  </dd>
                                </div>
                              ))}
                            </dl>
                          </div>
                        )}
                      </article>
                    ))}
                  </div>
                </section>
              ))}
            </div>
            )}
              </>
            )}

            {activeModule === 'qualidade' && (
              <div className="documents-module">
                <p className="selected-project-name">Auditoria de preenchimento, rastreabilidade, cronograma, medição e governança dos dados dos Issues no ACC.</p>
                <div className="section-actions" style={{ marginBottom: 12 }}>
                  <button className="ghost-button" type="button" onClick={async () => {
                    const report = await requestJson(`/api/projects/${encodeURIComponent(selectedProjectId)}/cronograma/issue-report`);
                    setQualityReportInfo(report.file || null);
                    const dashboard = buildQualityDashboard(issues || [], qualityData || null);
                    setQualityData(dashboard);
                    localStorage.setItem(`quality-cache:${selectedProjectId}`, JSON.stringify({ report: report.file, dashboard, processedAt: new Date().toISOString() }));
                  }}>Atualizar relatórios do ACC</button>
                  <button className="ghost-button" type="button" onClick={() => { const d = buildQualityDashboard(issues || [], qualityData || null); setQualityData(d); }}>Ler relatório mais recente</button>
                </div>
                <div className="module-tabs">{QUALITY_TABS.map((tab) => <button key={tab} type="button" className={`ghost-button ${qualityTab===tab?'active':''}`} onClick={() => setQualityTab(tab)}>{tab}</button>)}</div>
                {!!qualityData && <div className="impact-cards">
                  <article className="impact-card"><h4>Total de Issues analisados</h4><strong>{qualityData.cards.total}</strong></article>
                  <article className="impact-card"><h4>Issues abertos</h4><strong>{qualityData.cards.open}</strong></article>
                  <article className="impact-card"><h4>Issues fechados</h4><strong>{qualityData.cards.closed}</strong></article>
                  <article className="impact-card"><h4>Issues atrasados</h4><strong>{qualityData.cards.overdue}</strong></article>
                  <article className="impact-card"><h4>IQI geral</h4><strong>{qualityData.avgIqi.toFixed(1)}%</strong></article>
                  <article className="impact-card"><h4>Variação IQI</h4><strong>{qualityData.deltaIqi.toFixed(1)} p.p.</strong></article>
                </div>}
                {qualityReportInfo && <p className="selected-project-name">Relatório fonte: <strong>{qualityReportInfo.name}</strong> • versão {qualityReportInfo.version || 'n/d'} • atualizado em {formatDateTime(qualityReportInfo.updatedAt)} • pasta {qualityReportInfo.folderPath || 'n/d'}</p>}
              </div>
            )}

            {activeModule === 'schedule' && (
              <section className={`panel-card eap-module-card schedule-module-card ${scheduleViewMode === 'coordenacao' ? 'is-coordination-mode' : ''}`}>
                <div className="eap-module-header">
                  <div>
                    <p className="eyebrow">VISÃO GERAL</p>
                    <h2 className="eap-module-title">Visão geral do projeto</h2>
                    <p className="selected-project-name">Projeto selecionado: <strong>{selectedProject?.name || selectedProjectId || 'N/D'}</strong></p>
                    <p className="muted">Motor de cálculo sobre Issues do ACC para marcos, tramitação, riscos, dependências e avanço executivo.</p>
                  </div>
                </div>

                <section className="schedule-main-actions" aria-label="Ações principais da visão geral do projeto">
                  <button className="ghost-button" type="button" onClick={() => setActiveModule('')}>Módulos</button>
                  <button className="ghost-button" type="button" onClick={refreshCurrentModule} disabled={issuesLoading}>Atualizar</button>
                  <button className="ghost-button" type="button" onClick={() => setScheduleAdvancedFiltersOpen(true)}>Marcar feriados</button>
                  <button className="secondary-action" type="button" onClick={saveScheduleChanges} disabled={savingSchedule || (!schedulePendingEditCount && !scheduleFormulaSuggestions.length)}>
                    {savingSchedule ? 'Salvando...' : 'Salvar no CDE/ACC'}
                  </button>
                  <button className="ghost-button" type="button" onClick={exportSchedulePowerBiCsv} disabled={!schedulePlannerRows.length}>Exportar Power BI</button>
                  <button className="ghost-button" type="button" onClick={exportScheduleMsProjectCsv} disabled={!schedulePlannerRows.length}>Exportar MS Project</button>
                </section>

                <section className="schedule-kpi-strip" aria-label="Indicadores do cronograma">
                  {scheduleKpiCards.map((card) => (
                    <article key={card.id} className={`schedule-kpi-card is-${card.tone}`}>
                      <strong>{card.value}</strong>
                      <span>{card.label}</span>
                    </article>
                  ))}
                </section>

                <section className="schedule-filter-panel">
                  <div className="schedule-filter-title">
                    <div>
                      <p className="eyebrow">FILTROS</p>
                      <h3>Leitura do cronograma</h3>
                    </div>
                    <button type="button" className="ghost-button" onClick={() => setScheduleAdvancedFiltersOpen((current) => !current)}>
                      Filtros avançados
                    </button>
                  </div>
                  <div className="schedule-controls schedule-filter-controls">
                    <label>
                      Modo
                      <select value={scheduleViewMode} onChange={(event) => setScheduleViewMode(event.target.value)}>
                        {scheduleViewModes.map((mode) => <option key={mode.id} value={mode.id}>{mode.label}</option>)}
                      </select>
                    </label>
                    <label>
                      Agrupar por
                      <select value={scheduleGroupBy} onChange={(event) => setScheduleGroupBy(event.target.value)}>
                        <option value="codigoMarco">Código do Marco</option>
                        <option value="marcoContratual">Entrega vinculada</option>
                        <option value="eapVinculada">EAP Vinculada</option>
                        <option value="fase">Fase</option>
                        <option value="areaResponsavel">Área responsável</option>
                        <option value="disciplinaEnvolvida">Disciplina envolvida</option>
                        <option value="typeLabel">Tipo do Issue</option>
                        <option value="faseFluxo">Fase do Fluxo</option>
                        <option value="statusEntrega">Status da entrega</option>
                        <option value="prioridadeGestao">Prioridade de Gestão</option>
                      </select>
                    </label>
                    <label className="schedule-search">
                      Busca
                      <input value={scheduleSearch} onChange={(event) => setScheduleSearch(event.target.value)} placeholder="Título, ID, marco, documento, EAP..." />
                    </label>
                    <label>
                      Status
                      <select value={scheduleStatusFilter} onChange={(event) => setScheduleStatusFilter(event.target.value)}>
                        <option value="all">Todos</option>
                        {scheduleStatusOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Responsável
                      <select value={scheduleOwnerFilter} onChange={(event) => setScheduleOwnerFilter(event.target.value)}>
                        <option value="all">Todos</option>
                        {scheduleOwnerOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <label>
                      Marco
                      <select value={scheduleMarcoFilter} onChange={(event) => setScheduleMarcoFilter(event.target.value)}>
                        <option value="all">Todos</option>
                        {scheduleMarcoOptions.map((option) => <option key={option} value={option}>{option}</option>)}
                      </select>
                    </label>
                    <button type="button" className="ghost-button" onClick={() => { setScheduleSearch(''); setScheduleViewMode('executiva'); setScheduleStatusFilter('all'); setScheduleOwnerFilter('all'); setScheduleMarcoFilter('all'); }}>
                      Limpar filtros
                    </button>
                  </div>
                  {scheduleAdvancedFiltersOpen && (
                    <div className="schedule-advanced-panel">
                      <p><strong>Configurações rápidas:</strong> filtros avançados ficam preparados aqui para feriados, calendários e regras adicionais do cronograma.</p>
                      <p className="muted">A leitura principal continua compacta para deixar mais espaço para a planilha e para o Gantt.</p>
                    </div>
                  )}
                </section>

                <section className="schedule-top-grid">
                  <div className="schedule-command-panel">
                    <p className="eyebrow">MÓDULO DE VISUALIZAÇÃO</p>
                    <h3>Configuração da leitura</h3>
                    <div className="schedule-controls">
                      <label>
                        Modo de visualização
                        <select value={scheduleViewMode} onChange={(event) => setScheduleViewMode(event.target.value)}>
                          {scheduleViewModes.map((mode) => <option key={mode.id} value={mode.id}>{mode.label}</option>)}
                        </select>
                      </label>
                      <label>
                        Agrupar por
                        <select value={scheduleGroupBy} onChange={(event) => setScheduleGroupBy(event.target.value)}>
                          <option value="codigoMarco">Código do Marco</option>
                          <option value="marcoContratual">Entrega vinculada</option>
                          <option value="eapVinculada">EAP Vinculada</option>
                          <option value="fase">Fase</option>
                          <option value="areaResponsavel">Área responsável</option>
                          <option value="disciplinaEnvolvida">Disciplina envolvida</option>
                          <option value="typeLabel">Tipo do Issue</option>
                          <option value="faseFluxo">Fase do Fluxo</option>
                          <option value="statusEntrega">Status da entrega</option>
                          <option value="prioridadeGestao">Prioridade de Gestão</option>
                        </select>
                      </label>
                      <label className="schedule-search">
                        Busca
                        <input value={scheduleSearch} onChange={(event) => setScheduleSearch(event.target.value)} placeholder="Título, ID, marco, documento, EAP, predecessor..." />
                      </label>
                      <button type="button" className="ghost-button" onClick={() => { setScheduleSearch(''); setScheduleViewMode('executiva'); }}>Limpar filtros</button>
                    </div>
                  </div>

                  <div className="schedule-summary-panel">
                    <p className="eyebrow">QUADRO RESUMO</p>
                    <h3>Status executivo</h3>
                    <ul className="schedule-summary-list">
                      <li><span>Total de marcos</span><strong>{scheduleDashboard.totalMarcos}</strong></li>
                      <li><span>Issues vinculados</span><strong>{scheduleDashboard.totalVinculados}</strong></li>
                      <li><span>Marcos no prazo</span><strong>{scheduleDashboard.marcosNoPrazo}</strong></li>
                      <li className={scheduleDashboard.marcosAtrasados ? 'is-danger' : ''}><span>Marcos atrasados</span><strong>{scheduleDashboard.marcosAtrasados}</strong></li>
                      <li><span>Em desenvolvimento</span><strong>{scheduleDashboard.desenvolvimento}</strong></li>
                      <li><span>Emitidas ao cliente</span><strong>{scheduleDashboard.emitidas}</strong></li>
                      <li><span>Em análise cliente</span><strong>{scheduleDashboard.analiseCliente}</strong></li>
                      <li><span>Com comentários</span><strong>{scheduleDashboard.comentarios}</strong></li>
                      <li><span>Em revisão interna</span><strong>{scheduleDashboard.revisaoInterna}</strong></li>
                      <li><span>Aprovadas</span><strong>{scheduleDashboard.aprovadas}</strong></li>
                      <li><span>Pendências abertas</span><strong>{scheduleDashboard.pendenciasAbertas}</strong></li>
                      <li><span>Riscos abertos</span><strong>{scheduleDashboard.riscosAbertos}</strong></li>
                      <li className={scheduleDashboard.semMarco ? 'is-danger' : ''}><span>Sem marco</span><strong>{scheduleDashboard.semMarco}</strong></li>
                      <li className={scheduleDashboard.predecessorPendente ? 'is-danger' : ''}><span>Predecessor pendente</span><strong>{scheduleDashboard.predecessorPendente}</strong></li>
                      <li><span>Atualizações calculadas</span><strong>{scheduleDashboard.atualizacoesPendentes}</strong></li>
                      <li><span>Alterações manuais</span><strong>{scheduleDashboard.alteracoesPendentes}</strong></li>
                    </ul>
                  </div>
                </section>

                <section className="eap-gantt-panel schedule-gantt-panel">
                  <div className="cronograma-dashboard-header">
                    <div>
                      <h3>Gantt do projeto</h3>
                      <p className="muted">Barra clara: planejado. Barra verde: realizado. Marcadores: contrato, meta, retorno, aprovação e data atual.</p>
                    </div>
                    <div className="schedule-gantt-legend">
                      <span><i className="marker-blue" /> Contratual</span>
                      <span><i className="marker-green" /> Meta interna</span>
                      <span><i className="marker-orange" /> Retorno cliente</span>
                      <span><i className="marker-purple" /> Retorno real</span>
                      <span><i className="marker-black" /> Aprovação</span>
                    </div>
                  </div>
                  <div className="eap-gantt schedule-gantt">
                    <div className="eap-gantt-axis">
                      <span>Item</span>
                      <div>
                        {scheduleGantt.ticks.map((tick) => <small key={tick.toISOString()}>{formatDate(tick.toISOString())}</small>)}
                      </div>
                      <span>%</span>
                    </div>
                    <div className="eap-gantt-rows">
                      {scheduleGantt.items.map((row) => (
                        <div
                          key={row.id}
                          className={`eap-gantt-row schedule-gantt-row ${row.isMarco ? 'is-marco' : ''} ${row.maxDelay > 0 ? 'is-overdue-row' : ''} ${scheduleSelectedRowId === row.id ? 'is-selected' : ''}`}
                          onClick={() => selectScheduleRow(row.id)}
                          role="button"
                          tabIndex={0}
                          onKeyDown={(event) => { if (event.key === 'Enter' || event.key === ' ') selectScheduleRow(row.id); }}
                        >
                          <span className="eap-gantt-label" title={row.issue.title}>{row.plannerLevel ? '   ↳ ' : ''}{row.eapAuto || row.fields.codigoMarco || row.fields.eapVinculada || row.issue.displayId || row.id} - {row.issue.title}</span>
                          <div className="eap-gantt-track">
                            {scheduleGantt.todayLeft !== null && <i className="schedule-today-line" style={{ left: `${scheduleGantt.todayLeft}%` }} />}
                            <i className={`planned-bar ${row.maxDelay > 0 ? 'is-overdue' : ''}`} style={{ left: `${row.ganttLeft}%`, width: `${row.ganttWidth}%` }} />
                            {row.realStart && <i className={`real-bar ${row.completed ? 'is-done' : ''}`} style={{ left: `${row.realLeft}%`, width: `${row.realWidth}%` }}><b /></i>}
                            {row.markers.map((marker) => <em key={`${row.id}:${marker.key}`} className="schedule-date-marker" style={{ left: `${marker.left}%`, backgroundColor: marker.color }} title={marker.label} />)}
                          </div>
                          <strong>{Math.round(row.progress || 0)}%</strong>
                        </div>
                      ))}
                    </div>
                  </div>
                </section>

                {scheduleViewMode === 'formulas' && (
                  <section className="schedule-formula-panel">
                    <div className="cronograma-dashboard-header">
                      <div>
                        <h3>Prévia de atualizações calculadas</h3>
                        <p className="muted">Marque somente o que deseja gravar no ACC. Campos manuais importantes não são sobrescritos sem sua confirmação.</p>
                      </div>
                      <button type="button" className="ghost-button" onClick={() => selectAllScheduleFormulaSuggestions(false)}>Desmarcar tudo</button>
                    </div>
                    <div className="schedule-formula-list">
                      {scheduleFormulaSuggestions.map((suggestion) => (
                        <label key={suggestion.id} className={`schedule-formula-row ${suggestion.safeToOverwrite ? '' : 'needs-review'}`}>
                          <input type="checkbox" checked={!!scheduleFormulaSelections[suggestion.id]} onChange={(event) => toggleScheduleFormulaSuggestion(suggestion.id, event.target.checked)} />
                          <span><strong>{suggestion.issueTitle}</strong><small>{suggestion.label}</small></span>
                          <span className="formula-current">{suggestion.current || 'vazio'}</span>
                          <span className="formula-next">{suggestion.value}</span>
                          <small>{suggestion.reason}</small>
                        </label>
                      ))}
                      {!scheduleFormulaSuggestions.length && <p className="muted">Nenhuma atualização calculada pendente para os campos existentes no ACC.</p>}
                    </div>
                  </section>
                )}

                <section className="schedule-spreadsheet-panel">
                <div className="schedule-planner-toolbar" aria-label="Comandos da planilha da visão geral do projeto">
                  <span className="schedule-toolbar-title">Planilha do cronograma</span>
                  <span className={`schedule-autosave-badge is-${scheduleAutosaveState}`} title={scheduleAutosaveMessage || 'Salvamento automático semelhante ao Kanban'}>
                    {savingSchedule && scheduleAutosaveState === 'saving' ? 'Salvando...' : scheduleAutosaveState === 'error' ? 'Erro no autosave' : scheduleAutosaveState === 'saved' ? 'Salvo automaticamente' : scheduleAutosaveState === 'pending' ? 'Autosave pendente' : 'Autosave ativo'}
                  </span>
                  <button type="button" onClick={() => createScheduleIssue({
                    category: 'Gestão de Entregas',
                    type: 'Marco Contratual / Entrega',
                    defaultTitle: 'Novo marco contratual',
                    description: 'Marco criado pelo modulo Cronograma.',
                    isMarco: true,
                    usePlanningFields: true
                  })}>+ Novo marco</button>
                  <button type="button" onClick={() => createScheduleIssue({
                    category: 'Gestão de Entregas',
                    type: 'Pendência de Entrega',
                    defaultTitle: 'Nova pendência de entrega',
                    description: 'Pendencia criada pelo modulo Cronograma.',
                    usePlanningFields: true
                  })}>+ Nova tarefa</button>
                  <button type="button" onClick={() => createScheduleIssue({
                    category: 'Gestão de Entregas',
                    type: 'Comentário Cliente / Atendimento',
                    defaultTitle: 'Nova emissão / atendimento ao cliente',
                    description: 'Emissao ou atendimento criado pelo modulo Cronograma.',
                    usePlanningFields: true
                  })}>+ Emissão</button>
                  <button type="button" onClick={() => createScheduleIssue({
                    category: 'Gestão de Entregas',
                    type: 'Restrição / Risco',
                    defaultTitle: 'Nova restrição ou risco',
                    description: 'Restricao ou risco criado pelo modulo Cronograma.',
                    usePlanningFields: true
                  })}>+ Restrição/Risco</button>
                  <button type="button" onClick={() => createScheduleIssue({
                    category: 'Interface e Coordenação Multidisciplinar',
                    type: 'Solicitação de Informação',
                    defaultTitle: 'Nova solicitação de informação',
                    description: 'Solicitacao de informacao criada pelo modulo Cronograma.',
                    usePlanningFields: false
                  })}>+ Solicitação de Informação</button>
                  <span className="schedule-toolbar-divider" aria-hidden="true" />
                  <button type="button" onClick={indentSelectedScheduleRow} disabled={!scheduleSelectedRowId}>Tornar filho</button>
                  <button type="button" onClick={outdentSelectedScheduleRow} disabled={!scheduleSelectedRowId}>Voltar nível</button>
                  <button type="button" onClick={() => moveSelectedScheduleRow(-1)} disabled={!scheduleSelectedRowId}>Mover acima</button>
                  <button type="button" onClick={() => moveSelectedScheduleRow(1)} disabled={!scheduleSelectedRowId}>Mover abaixo</button>
                  <button
                    type="button"
                    className="deliverable-title-toggle"
                    onClick={toggleSelectedScheduleDeliverableTitle}
                    disabled={!scheduleSelectedRowId}
                  >
                    {scheduleSelectedRowId && scheduleRowHighlights[scheduleSelectedRowId] === 'deliverableTitle' ? 'Remover Título' : 'Marcar Título'}
                  </button>
                </div>

                <div className="table-shell cronograma-table-shell eap-table-wrapper schedule-table-wrapper">
                  <table className="eap-table schedule-table">
                    <thead>
                      <tr>
                        {scheduleColumns.map((column) => (
                          <th key={column} className={getScheduleColumnClassName(column)}>
                            {getScheduleColumnHeaderParts(column).map((part) => <span key={part}>{part}</span>)}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {schedulePlannerRows.map((row) => (
                        <tr
                          key={row.id}
                          className={`${row.isMarco ? 'marco-row' : ''} ${scheduleRowHighlights[row.id] === 'deliverableTitle' ? 'deliverable-title-row' : ''} ${row.linkWarning ? 'schedule-warning-row' : ''} ${getScheduleDependencyState(row).blocked ? 'schedule-blocked-row' : ''} ${getScheduleDependencyState(row).released ? 'schedule-released-row' : ''} ${scheduleSelectedRowId === row.id ? 'selected-schedule-row' : ''}`}
                          onClick={() => selectScheduleRow(row.id)}
                        >
                          {scheduleColumns.map((column) => <td key={`${row.id}:${column}`} className={getScheduleColumnClassName(column)}>{renderScheduleCell(row, column)}</td>)}
                        </tr>
                      ))}
                      {!schedulePlannerRows.length && (
                        <tr><td colSpan={scheduleColumns.length}>Nenhum Issue encontrado para os filtros atuais.</td></tr>
                      )}
                    </tbody>
                  </table>
                </div>

                <aside className="schedule-detail-panel" aria-live="polite">
                  {selectedScheduleRow ? (
                    <>
                      <p className="eyebrow">ITEM SELECIONADO</p>
                      <h3>{selectedScheduleRow.issue?.title || getScheduleCellValue(selectedScheduleRow, 'title')}</h3>
                      <div className="schedule-detail-meta">
                        <span>{selectedScheduleRow.eapAuto || 'Sem EAP'}</span>
                        <span>{selectedScheduleRow.typeLabel || 'Issue'}</span>
                      </div>
                      <div className="schedule-detail-grid">
                        {renderScheduleNativeStatusEditor(selectedScheduleRow)}
                        {renderScheduleIssueCategoryEditor(selectedScheduleRow)}
                        {renderScheduleIssueTypeEditor(selectedScheduleRow)}
                        {renderScheduleAssigneeEditor(selectedScheduleRow)}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'disciplinaEnvolvida', 'Disciplina')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'statusEntrega', 'Status da entrega')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'statusCliente', 'Status Cliente')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'inicioPlanejado', 'Início planejado')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'terminoPlanejado', 'Término planejado')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'diasPrevistosAtividade', 'Dias previstos')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'codigoMarco', 'Código do Marco')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'predecessor', 'Predecessor')}
                        {renderScheduleDetailEditor(selectedScheduleRow, 'prioridadeGestao', 'Prioridade')}
                      </div>
                      <label className="schedule-detail-field is-wide">
                        Ação necessária
                        <textarea
                          value={getScheduleCellValue(selectedScheduleRow, 'acaoNecessaria')}
                          onChange={(event) => editScheduleCell(selectedScheduleRow.id, 'acaoNecessaria', event.target.value)}
                          rows={4}
                          placeholder="Registrar ação necessária"
                        />
                      </label>
                      <div className="schedule-detail-actions">
                        {(selectedScheduleRow.issue?.webUrl || selectedScheduleRow.issue?.url) && (
                          <a className="ghost-button" href={selectedScheduleRow.issue.webUrl || selectedScheduleRow.issue.url} target="_blank" rel="noreferrer">Abrir no ACC</a>
                        )}
                        <button
                          type="button"
                          className="secondary-action"
                          onClick={saveScheduleChanges}
                          disabled={savingSchedule || (!schedulePendingEditCount && !scheduleFormulaSuggestions.length)}
                        >
                          {savingSchedule ? 'Salvando...' : 'Salvar no CDE/ACC'}
                        </button>
                      </div>
                    </>
                  ) : (
                    <div className="schedule-detail-empty">
                      <p className="eyebrow">DETALHES</p>
                      <h3>Selecione uma linha</h3>
                      <p className="muted">Clique em uma tarefa, marco ou barra do Gantt para ver e editar os principais dados do issue.</p>
                    </div>
                  )}
                </aside>
                </section>

                <section className="schedule-field-status">
                  <p className="eyebrow">Campos calculados sem destino no ACC</p>
                  <p>{scheduleMissingCalculatedTargets.length ? scheduleMissingCalculatedTargets.join(', ') : 'Todos os principais campos calculados têm campo correspondente encontrado no ACC.'}</p>
                  {scheduleMissingStructuralTargets.length > 0 && (
                    <p className="muted">Campos estruturais não encontrados para salvar EAP/filho: {scheduleMissingStructuralTargets.join(', ')}.</p>
                  )}
                </section>
              </section>
            )}

            {activeModule === 'cronograma-im' && (
              <section className="panel-card eap-module-card">
                <div className="eap-module-header">
                  <div>
                    <p className="eyebrow">EAP</p>
                    <h2 className="eap-module-title">EAP</h2>
                    <p className="selected-project-name">Projeto selecionado: <strong>{selectedProject?.name || selectedProjectId || 'N/D'}</strong></p>
                    <p className="muted">Visão consolidada e somente leitura dos issues estruturados por EAP, marco, prazo e execução.</p>
                  </div>
                  <div className="eap-module-actions">
                    <button className="ghost-button" type="button" onClick={() => setActiveModule('')}>Módulos</button>
                    <button className="ghost-button" type="button" onClick={refreshCurrentModule} disabled={issuesLoading || issueReportLoading}>Atualizar</button>
                    <button type="button" className="secondary-action" onClick={exportCronogramaMsProjectCsv} disabled={!cronogramaRows.length}>
                      Exportar CSV MS Project
                    </button>
                    <button type="button" className="secondary-action" onClick={exportCronogramaPowerBiCsv} disabled={!cronogramaRows.length}>
                      Exportar Power BI
                    </button>
                  </div>
                </div>
                {issueReportStatus && <p className="muted cronograma-status-text">{issueReportStatus}</p>}
                {issueReportInfo && <p className="muted">Relatório usado: {issueReportInfo.name} (v{issueReportInfo.version || 'N/D'}) — pasta: {issueReportInfo.folderPath}</p>}

                <section className="eap-top-summary">
                  <article className="eap-summary-card">
                    <div>
                      <p className="eyebrow">Execução geral</p>
                      <strong>{cronogramaProjectExecution}%</strong>
                      <span>{cronogramaRows.filter((row) => row.concluded).length} concluídos de {cronogramaRows.length} issues estruturados.</span>
                    </div>
                    <div className="eap-donut small" style={{ '--pct': `${cronogramaProjectExecution}%` }}>
                      <span>{cronogramaProjectExecution}%</span>
                    </div>
                  </article>
                  <article className="eap-summary-card">
                    <div>
                      <p className="eyebrow">Próximos vencimentos</p>
                      <strong>{eapUpcomingRows.length}</strong>
                      <span>Atividades vencendo nos próximos 7 dias.</span>
                      {eapUpcomingRows[0] && <small>{eapUpcomingRows[0].row.codigoMarco || eapUpcomingRows[0].row.marcoContratual} • {formatDate(eapUpcomingRows[0].target)}</small>}
                    </div>
                  </article>
                  <article className="eap-summary-card attention">
                    <div>
                      <p className="eyebrow">Pontos de atenção</p>
                      <strong>{eapAttentionSummary.atrasados + eapAttentionSummary.bloqueados + eapAttentionSummary.semResponsavel}</strong>
                      <span>Atrasos, bloqueios ou itens sem responsável.</span>
                      <div className="eap-inline-tags">
                        <span>{eapAttentionSummary.atrasados} atrasados</span>
                        <span>{eapAttentionSummary.bloqueados} bloqueados</span>
                        <span>{eapAttentionSummary.semResponsavel} sem responsável</span>
                      </div>
                    </div>
                  </article>
                </section>

                <div className="metric-grid eap-metric-grid">
                  <div><span>Total de Issues</span><strong>{cronogramaRows.length}</strong></div>
                  <div><span>Concluídos</span><strong>{cronogramaRows.filter((row) => row.concluded).length}</strong></div>
                  <div><span>Atrasados</span><strong>{cronogramaRows.filter((row) => row.overdue).length}</strong></div>
                  <div><span>Bloqueados</span><strong>{cronogramaRows.filter((row) => row.blocked).length}</strong></div>
                  <div><span>Sem responsável</span><strong>{cronogramaRows.filter((row) => !row.atribuidoA || normalizeText(row.atribuidoA).includes('sem atribu')).length}</strong></div>
                  <div><span>Com predecessor</span><strong>{cronogramaDashboard.predecessorStats.predecessorValido}</strong></div>
                  <div><span>Inconsistências</span><strong>{cronogramaRows.filter((row) => row.predecessoresTexto.includes('ID não encontrado')).length}</strong></div>
                </div>

                <div className="eap-view-toolbar">
                  <div className="eap-toolbar-group">
                    <span className="eap-toolbar-label">Painel</span>
                    <div className="eap-pill-group">
                      <button type="button" className={`eap-view-pill ${eapPanelMode === 'dashboard' ? 'is-active' : ''}`} onClick={() => setEapPanelMode('dashboard')}>Dashboard executivo</button>
                      <button type="button" className={`eap-view-pill ${eapPanelMode === 'estrutura' ? 'is-active' : ''}`} onClick={() => setEapPanelMode('estrutura')}>Gestão por entregas</button>
                    </div>
                  </div>
                  <div className="eap-toolbar-group">
                    <span className="eap-toolbar-label">Visualização da gestão por entregas</span>
                    <div className="eap-pill-group">
                      {[
                        { key: 'entregas', label: 'Entregas' },
                        { key: 'cronograma', label: 'Cronograma' },
                        { key: 'marcos', label: 'Marcos' },
                        { key: 'dependencias', label: 'Dependências' }
                      ].map((item) => (
                        <button
                          key={item.key}
                          type="button"
                          className={`eap-view-pill ${eapStructureView === item.key ? 'is-active' : ''}`}
                          onClick={() => { setEapPanelMode('estrutura'); setEapStructureView(item.key); }}
                        >
                          {item.label}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>

                {eapPanelMode === 'dashboard' ? (
                  <>
                    <section className="cronograma-dashboard">
                      <div className="cronograma-dashboard-header">
                        <div>
                          <h3>Dashboard de Issues</h3>
                          <p className="muted">Visão gerencial dos issues por status, prazo, marco, tipo e execução.</p>
                        </div>
                        <div className="cronograma-dashboard-filter">
                          {cronogramaDashboardFilter && <span>Filtro ativo: {cronogramaDashboardFilter.label}</span>}
                          <button type="button" className="ghost-button" onClick={() => setCronogramaDashboardFilter(null)} disabled={!cronogramaDashboardFilter}>Limpar filtros</button>
                        </div>
                      </div>
                      {(() => {
                        const total = Math.max(1, cronogramaDashboard.total || cronogramaRows.length);
                        const concluidos = cronogramaRows.filter((row) => row.concluded).length;
                        const atrasados = cronogramaRows.filter((row) => row.overdue).length;
                        const bloqueados = cronogramaRows.filter((row) => row.blocked).length;
                        const abertos = Math.max(0, cronogramaRows.length - concluidos);
                        const prazoResumo = cronogramaDashboard.prazoData.slice(0, 4);
                        const tipoResumo = cronogramaDashboard.tipoData.slice(0, 5);
                        return (
                          <div className="eap-dashboard-showcase">
                            <article className="eap-donut-card">
                              <div className="eap-donut" style={{ '--pct': `${cronogramaProjectExecution}%` }}>
                                <span>{cronogramaProjectExecution}%</span>
                              </div>
                              <div>
                                <p className="eyebrow">EXECUÇÃO GERAL</p>
                                <h4>Panorama da EAP</h4>
                                <p className="muted">{concluidos} concluídos de {cronogramaRows.length} issues lidos.</p>
                              </div>
                            </article>
                            <article className="eap-chart-card">
                              <div className="eap-chart-title">
                                <h4>Status operacional</h4>
                                <span>{cronogramaRows.length} issues</span>
                              </div>
                              <div className="eap-status-strip">
                                {[
                                  { label: 'Concluídos', value: concluidos, className: 'is-done' },
                                  { label: 'Abertos', value: abertos, className: 'is-open' },
                                  { label: 'Atrasados', value: atrasados, className: 'is-late' },
                                  { label: 'Bloqueados', value: bloqueados, className: 'is-blocked' },
                                ].map((item) => (
                                  <button key={item.label} type="button" className={`eap-status-stat ${item.className}`} onClick={() => item.label === 'Atrasados' ? setCronogramaDashboardFilter({ kind: 'prazo', value: 'Atrasados', label: 'Prazo • Atrasados' }) : null}>
                                    <strong>{item.value}</strong>
                                    <span>{item.label}</span>
                                  </button>
                                ))}
                              </div>
                            </article>
                            <article className="eap-chart-card">
                              <div className="eap-chart-title">
                                <h4>Prazo e tipos</h4>
                                <span>Resumo visual</span>
                              </div>
                              <div className="eap-mini-bars">
                                {prazoResumo.map((item) => (
                                  <button key={item.name} type="button" className="eap-mini-bar" onClick={() => setCronogramaDashboardFilter({ kind: 'prazo', value: item.name, label: `Prazo • ${item.name}` })}>
                                    <span>{item.name}</span>
                                    <div><i style={{ width: `${(item.value / total) * 100}%`, background: item.color }} /></div>
                                    <strong>{item.value}</strong>
                                  </button>
                                ))}
                                {tipoResumo.map((item, index) => (
                                  <button key={item.name} type="button" className="eap-mini-bar compact" onClick={() => setCronogramaDashboardFilter({ kind: 'tipo', value: item.name, label: `Tipo • ${item.name}` })}>
                                    <span>{item.name}</span>
                                    <div><i style={{ width: `${(item.value / total) * 100}%`, background: ['#0f7c90', '#1e4f7a', '#64748b', '#c38b2f', '#d97373'][index % 5] }} /></div>
                                    <strong>{item.value}</strong>
                                  </button>
                                ))}
                              </div>
                            </article>
                          </div>
                        );
                      })()}
                      <div className="dashboard-grid two-cols">
                        <article className="dashboard-card"><h4>Issues por Status</h4>{cronogramaDashboard.statusData.map((item) => <button key={item.name} type="button" className="bar-row" title={`${item.name}: ${item.value} (${Math.round((item.value / Math.max(1, cronogramaDashboard.total)) * 100)}%)`} onClick={() => setCronogramaDashboardFilter({ kind: 'status', value: item.name, label: `Status • ${item.name}` })}><span>{item.name}</span><div><i style={{ width: `${(item.value / Math.max(1, cronogramaDashboard.total)) * 100}%` }} /></div><strong>{item.value}</strong></button>)}</article>
                        <article className="dashboard-card"><h4>Issues por Tipo</h4><div className="donut-list">{cronogramaDashboard.tipoData.map((item, index) => <button key={item.name} type="button" className="legend-item" onClick={() => setCronogramaDashboardFilter({ kind: 'tipo', value: item.name, label: `Tipo • ${item.name}` })}><i style={{ backgroundColor: ['#0f7c90', '#1e4f7a', '#64748b', '#c38b2f', '#d97373', '#b8c0cc'][index % 6] }} /><span>{item.name}</span><strong>{item.value}</strong></button>)}</div></article>
                      </div>
                      <div className="dashboard-grid two-cols">
                        <article className="dashboard-card"><h4>Issues por Marco</h4>{cronogramaDashboard.marcoData.map((item) => <button key={item.name} type="button" className="bar-row" onClick={() => setCronogramaDashboardFilter({ kind: 'marco', value: item.name, label: `Marco • ${item.name}` })}><span>{item.name}</span><div><i style={{ width: `${(item.value / Math.max(1, cronogramaDashboard.total)) * 100}%`, background: '#1e4f7a' }} /></div><strong>{item.value}</strong></button>)}</article>
                        <article className="dashboard-card"><h4>Distribuição por Prazo</h4><div className="donut-list">{cronogramaDashboard.prazoData.map((item) => <button key={item.name} type="button" className="legend-item" onClick={() => setCronogramaDashboardFilter({ kind: 'prazo', value: item.name, label: `Prazo • ${item.name}` })}><i style={{ backgroundColor: item.color }} /><span>{item.name}</span><strong>{item.value}</strong></button>)}</div></article>
                      </div>
                    </section>
                    <div className="table-shell cronograma-table-shell eap-table-wrapper eap-readonly-wrapper">
                      <table className="eap-table eap-table-readonly">
                        <thead>
                          <tr>
                            <th>{renderCronogramaSortHeader('eapVinculada', 'EAP')}</th>
                            <th>{renderCronogramaSortHeader('titulo', 'Título')}</th>
                            <th>{renderCronogramaSortHeader('inicioPrevisto', 'Início Planejado')}</th>
                            <th>{renderCronogramaSortHeader('terminoPrevisto', 'Término Planejado')}</th>
                            <th>{renderCronogramaSortHeader('inicioReal', 'Início real')}</th>
                            <th>{renderCronogramaSortHeader('terminoReal', 'Término real')}</th>
                            <th>{renderCronogramaSortHeader('atribuidoA', 'Atribuído a')}</th>
                            <th>{renderCronogramaSortHeader('status', 'Status')}</th>
                            <th>{renderCronogramaSortHeader('delayDays', 'Dias de atraso')}</th>
                            <th>{renderCronogramaSortHeader('executado', 'Executado')}</th>
                            <th>{renderCronogramaSortHeader('marcoContratual', 'Marco Contratual')}</th>
                          </tr>
                        </thead>
                        <tbody>
                          {cronogramaRowsSorted.map((row) => {
                            const statusLabel = row.statusValue || row.issue.status || 'Sem status';
                            const statusClass = row.concluded ? 'is-done' : row.blocked ? 'is-blocked' : row.overdue ? 'is-late' : 'is-open';
                            const hasDelay = Number(row.delayDays || 0) > 0 && !row.concluded;
                            return (
                              <tr key={row.issue.id} className={`${row.isMarcoContratual ? 'marco-row' : ''} ${hasDelay ? 'is-overdue-row' : ''}`}>
                                <td><span className="eap-id-pill">{row.eapVinculada || '-'}</span></td>
                                <td className="eap-title-cell" title={row.titulo}>{renderEapLinkedIssueTitle(row)}</td>
                                <td>{formatDate(row.inicioPrevisto) || '-'}</td>
                                <td>{formatDate(row.terminoPrevisto) || '-'}</td>
                                <td>{formatDate(row.inicioReal) || '-'}</td>
                                <td>{formatDate(row.terminoReal) || '-'}</td>
                                <td>{row.atribuidoA || 'Sem atribuição'}</td>
                                <td><span className={`eap-status-pill ${statusClass}`}>{statusLabel}</span></td>
                                <td><span className={`eap-delay-pill ${hasDelay ? 'is-late' : ''}`}>{row.delayDays === null || row.delayDays === undefined ? 'N/C' : `${row.delayDays} d`}</span></td>
                                <td><div className="execution-cell"><span>{row.executado || 0}%</span><div className="execution-bar"><i style={{ width: `${row.executado || 0}%` }} /></div></div></td>
                                <td>{row.marcoContratual || '-'}</td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    </div>
                  </>
                ) : (
                  <section className="eap-structure-workspace">
                    <div className="eap-structure-header">
                      <div className="eap-quick-filters">
                        {[
                          { key: 'all', label: `Todas ${eapQuickFilterCounts.all}` },
                          { key: 'atrasadas', label: `Atrasadas ${eapQuickFilterCounts.atrasadas}` },
                          { key: 'bloqueadas', label: `Bloqueadas ${eapQuickFilterCounts.bloqueadas}` },
                          { key: 'semResponsavel', label: `Sem responsável ${eapQuickFilterCounts.semResponsavel}` },
                          { key: 'vencem7', label: `Vencem em 7 dias ${eapQuickFilterCounts.vencem7}` },
                          { key: 'semPredecessor', label: `Sem predecessor ${eapQuickFilterCounts.semPredecessor}` }
                        ].map((item) => (
                          <button
                            key={item.key}
                            type="button"
                            className={`eap-filter-chip ${eapQuickFilter === item.key ? 'is-active' : ''}`}
                            onClick={() => setEapQuickFilter(item.key)}
                          >
                            {item.label}
                          </button>
                        ))}
                      </div>
                      <button type="button" className="ghost-button" onClick={() => setEapQuickFilter('all')}>Limpar filtros</button>
                    </div>

                    <div className="eap-structure-layout">
                      <aside className="eap-delivery-sidebar">
                        <div className="eap-delivery-sidebar-head">
                          <p className="eyebrow">ENTREGAS PRINCIPAIS</p>
                          <h3>{eapDeliveryData.groups.length} entregas</h3>
                          <p className="muted">Agrupamento pelas issues cujo título inicia com “Entrega -”.</p>
                        </div>
                        <div className="eap-delivery-list">
                          <button
                            type="button"
                            className={`eap-delivery-card eap-delivery-all ${eapSelectedDeliveryId === 'all' ? 'is-active' : ''}`}
                            onClick={() => { setEapSelectedDeliveryId('all'); setEapSelectedRowId(''); }}
                          >
                            <span className="eap-delivery-code">TODAS</span>
                            <strong>Ver todas as entregas e atividades</strong>
                            <span className="eap-delivery-meta">{cronogramaRowsSorted.length} issues estruturados • visão geral do projeto</span>
                            <div className="eap-delivery-progress"><i style={{ width: `${cronogramaProjectExecution}%` }} /></div>
                            <div className="eap-delivery-stats">
                              <span>{cronogramaProjectExecution}%</span>
                              <span>{cronogramaRows.filter((row) => row.overdue && !row.concluded).length} atrasados</span>
                              <span>{cronogramaRows.filter((row) => row.blocked).length} bloqueados</span>
                            </div>
                          </button>
                          {eapDeliveryData.groups.map((group) => (
                            <button
                              key={group.id}
                              type="button"
                              className={`eap-delivery-card ${eapSelectedDeliveryId === group.id ? 'is-active' : ''}`}
                              onClick={() => { setEapSelectedDeliveryId(group.id); setEapSelectedRowId(group.delivery.issue.id); }}
                            >
                              <span className="eap-delivery-code">{group.eap || '-'}</span>
                              <strong>{group.title.replace(/^Entrega\s*-\s*/i, '')}</strong>
                              <span className="eap-delivery-meta">{group.total} itens vinculados • {group.marco || 'Sem marco'}</span>
                              <div className="eap-delivery-progress"><i style={{ width: `${group.progress}%` }} /></div>
                              <div className="eap-delivery-stats">
                                <span>{group.progress}%</span>
                                <span>{group.overdue} atrasados</span>
                                <span>{group.blocked} bloqueados</span>
                              </div>
                            </button>
                          ))}
                        </div>
                      </aside>

                      <div className="eap-structure-main">
                        <div className="eap-structure-panel">
                          <div className="eap-panel-heading">
                            <div>
                              <p className="eyebrow">{eapStructureView === 'marcos' ? 'MARCOS' : eapStructureView.toUpperCase()}</p>
                              <h3>
                                {eapStructureView === 'entregas' && (eapSelectedDelivery ? eapSelectedDelivery.title : 'Todas as entregas e atividades')}
                                {eapStructureView === 'cronograma' && (eapSelectedDelivery ? 'Cronograma estruturado da entrega' : 'Cronograma geral do projeto')}
                                {eapStructureView === 'marcos' && 'Resumo por marco contratual'}
                                {eapStructureView === 'dependencias' && 'Dependências e bloqueios'}
                              </h3>
                              <p className="muted">
                                {eapStructureView === 'entregas' && 'Visualização das entregas e atividades vinculadas. Para voltar ao projeto completo, selecione “Todas” na lista lateral.'}
                                {eapStructureView === 'cronograma' && 'Tabela ampliada para leitura do prazo, responsável, status e execução.'}
                                {eapStructureView === 'marcos' && 'Consolidação das entregas por marco contratual.'}
                                {eapStructureView === 'dependencias' && 'Foco nas atividades com predecessor, bloqueios e possíveis inconsistências.'}
                              </p>
                            </div>
                          </div>

                          {eapStructureView === 'marcos' ? (
                            <div className="eap-marco-grid">
                              {eapMarcoSummary.map((item) => (
                                <article key={item.name} className="eap-marco-card">
                                  <div className="eap-marco-top">
                                    <strong>{item.name}</strong>
                                    <span>{item.progress}%</span>
                                  </div>
                                  <div className="eap-delivery-progress"><i style={{ width: `${item.progress}%` }} /></div>
                                  <div className="eap-marco-stats">
                                    <span>{item.total} itens</span>
                                    <span>{item.concluidos} concluídos</span>
                                    <span>{item.atrasados} atrasados</span>
                                    <span>{item.pendentes} pendentes</span>
                                  </div>
                                </article>
                              ))}
                            </div>
                          ) : (
                            <>
                              {eapStructureView === 'cronograma' && (
                                <div className="eap-gantt-board">
                                  <div className="eap-gantt-head">
                                    <div>
                                      <p className="eyebrow">LINHA DO TEMPO</p>
                                      <h4>Cronograma visual</h4>
                                    </div>
                                    <span>{eapTimelineData.items.length} itens na visualização</span>
                                  </div>
                                  <div className="eap-gantt-ticks">
                                    <span />
                                    <div>
                                      {eapTimelineData.ticks.map((tick) => <i key={tick.toISOString()}>{formatDate(tick)}</i>)}
                                    </div>
                                  </div>
                                  <div className="eap-gantt-list">
                                    {eapTimelineData.items.map((item) => (
                                      <button key={item.issue.id} type="button" className={`eap-gantt-row ${eapSelectedRow?.issue.id === item.issue.id ? 'is-selected' : ''}`} onClick={() => setEapSelectedRowId(item.issue.id)}>
                                        <span className="eap-gantt-label"><strong>{item.eapVinculada || '-'}</strong>{item.titulo}</span>
                                        <span className="eap-gantt-lane">
                                          <i className="eap-gantt-today" style={{ left: `${eapTimelineData.todayLeft}%` }}><b>Hoje</b></i>
                                          <i className={`eap-gantt-bar ${item.overdue && !item.concluded ? 'is-late' : item.concluded ? 'is-done' : ''}`} style={{ left: `${item.ganttLeft}%`, width: `${item.ganttWidth}%` }}><b>{item.executado || 0}%</b></i>
                                          {item.realWidth > 0 && <i className="eap-gantt-real" style={{ left: `${item.realLeft}%`, width: `${item.realWidth}%` }} />}
                                        </span>
                                      </button>
                                    ))}
                                  </div>
                                </div>
                              )}
                              <div className="table-shell cronograma-table-shell eap-wide-table-wrapper">
                              <table className="eap-table eap-structure-table">
                                <thead>
                                  <tr>
                                    <th>EAP</th>
                                    <th>Título</th>
                                    <th>Responsável</th>
                                    <th>Início planejado</th>
                                    <th>Término planejado</th>
                                    <th>Início real</th>
                                    <th>Término real</th>
                                    <th>Status</th>
                                    <th>Atraso</th>
                                    <th>Execução</th>
                                    <th>Marco Contratual</th>
                                  </tr>
                                </thead>
                                <tbody>
                                  {eapViewRows.map((row) => {
                                    const statusClass = row.concluded ? 'is-done' : row.blocked ? 'is-blocked' : row.overdue ? 'is-late' : 'is-open';
                                    const hasDelay = Number(row.delayDays || 0) > 0 && !row.concluded;
                                    return (
                                      <tr
                                        key={row.issue.id}
                                        className={`${eapSelectedRow?.issue.id === row.issue.id ? 'is-selected-row' : ''} ${row.isMarcoContratual ? 'marco-row' : ''}`}
                                        onClick={() => setEapSelectedRowId(row.issue.id)}
                                      >
                                        <td><span className="eap-id-pill">{row.eapVinculada || '-'}</span></td>
                                        <td className={`eap-title-cell ${normalizeText(row.titulo).startsWith('entrega -') ? 'marco-title' : ''}`}>{renderEapLinkedIssueTitle(row)}</td>
                                        <td>{row.atribuidoA || 'Sem atribuição'}</td>
                                        <td>{formatDate(row.inicioPrevisto) || '-'}</td>
                                        <td>{formatDate(row.terminoPrevisto) || '-'}</td>
                                        <td>{formatDate(row.inicioReal) || '-'}</td>
                                        <td>{formatDate(row.terminoReal) || '-'}</td>
                                        <td><span className={`eap-status-pill ${statusClass}`}>{row.statusValue || row.issue.status || 'Sem status'}</span></td>
                                        <td><span className={`eap-delay-pill ${hasDelay ? 'is-late' : ''}`}>{row.delayDays === null || row.delayDays === undefined ? 'N/C' : `${row.delayDays} d`}</span></td>
                                        <td><div className="execution-cell"><span>{row.executado || 0}%</span><div className="execution-bar"><i style={{ width: `${row.executado || 0}%` }} /></div></div></td>
                                        <td>{row.marcoContratual || '-'}</td>
                                      </tr>
                                    );
                                  })}
                                </tbody>
                              </table>
                              </div>
                            </>
                          )}
                        </div>
                      </div>
                    </div>

                    {eapSelectedRow && eapStructureView !== 'marcos' && (
                      <section className="eap-detail-bottom">
                        <div className="eap-panel-heading">
                          <div>
                            <p className="eyebrow">DETALHE DO ITEM</p>
                            <h3>{eapSelectedRow.titulo}</h3>
                          </div>
                        </div>
                        <div className="eap-detail-grid">
                          <div><span>EAP</span><strong>{eapSelectedRow.eapVinculada || '-'}</strong></div>
                          <div><span>Marco</span><strong>{eapSelectedRow.marcoContratual || '-'}</strong></div>
                          <div><span>Responsável</span><strong>{eapSelectedRow.atribuidoA || 'Sem atribuição'}</strong></div>
                          <div><span>Início planejado</span><strong>{formatDate(eapSelectedRow.inicioPrevisto) || '-'}</strong></div>
                          <div><span>Término planejado</span><strong>{formatDate(eapSelectedRow.terminoPrevisto) || '-'}</strong></div>
                          <div><span>Início real</span><strong>{formatDate(eapSelectedRow.inicioReal) || '-'}</strong></div>
                          <div><span>Término real</span><strong>{formatDate(eapSelectedRow.terminoReal) || '-'}</strong></div>
                          <div><span>Data contratual</span><strong>{formatDate(eapSelectedRow.dataContratual) || '-'}</strong></div>
                          <div><span>Status</span><strong>{eapSelectedRow.statusValue || eapSelectedRow.issue.status || 'Sem status'}</strong></div>
                          <div><span>Atraso</span><strong>{eapSelectedRow.delayDays === null || eapSelectedRow.delayDays === undefined ? 'N/C' : `${eapSelectedRow.delayDays} dias`}</strong></div>
                          <div><span>Execução</span><strong>{eapSelectedRow.executado || 0}%</strong></div>
                        </div>
                      </section>
                    )}
                  </section>
                )}
              </section>
            )}
          </section>
            )}
          </>
        )}

        {loading && <p className="status">Verificando sessão Autodesk...</p>}
        {error && (
          <div className="app-alert-backdrop" role="presentation">
            <div
              className="app-alert-window"
              role="alertdialog"
              aria-modal="true"
              aria-labelledby="app-alert-title"
              aria-describedby="app-alert-message"
            >
              <div className="app-alert-titlebar">
                <div>
                  <p className="eyebrow">Alerta do Central G5</p>
                  <h3 id="app-alert-title">Atenção: ação não concluída</h3>
                </div>
                <button
                  className="app-alert-close"
                  type="button"
                  aria-label="Fechar aviso"
                  onClick={() => setError('')}
                >
                  ×
                </button>
              </div>

              <div className="app-alert-body">
                <p id="app-alert-message">{error}</p>
              </div>

              <div className="app-alert-actions">
                <button className="secondary-button" type="button" onClick={() => window.location.reload()}>
                  Recarregar tela
                </button>
                <button className="primary-button" type="button" autoFocus onClick={() => setError('')}>
                  Entendi
                </button>
              </div>
            </div>
          </div>
        )}
      </section>
    </main>
  );
}

createRoot(document.getElementById('root')).render(
  <AppErrorBoundary>
    <App />
  </AppErrorBoundary>
);
