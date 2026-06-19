import crypto from 'node:crypto';
import path from 'node:path';
import { fileURLToPath } from 'node:url';
import cors from 'cors';
import dotenv from 'dotenv';
import express from 'express';
import session from 'express-session';
import ExcelJS from 'exceljs';
import * as XLSX from 'xlsx';

dotenv.config();

const app = express();
const port = Number(process.env.PORT || 3000);
const frontendUrl = process.env.FRONTEND_URL || 'http://localhost:5173';
const callbackUrl = process.env.APS_CALLBACK_URL;
const sessionMaxAgeMs = Number(process.env.SESSION_MAX_AGE_MS || 1000 * 60 * 60 * 24 * 14);

const apsAuthorizeUrl = 'https://developer.api.autodesk.com/authentication/v2/authorize';
const apsTokenUrl = 'https://developer.api.autodesk.com/authentication/v2/token';
const apsApiBaseUrl = 'https://developer.api.autodesk.com';
const apsScopes = String(process.env.APS_SCOPES || 'data:read data:write data:create account:read').split(/\s+/).filter(Boolean);
const documentComparisonCache = new Map();
const documentComparisonCacheTtlMs = 30 * 60 * 1000;

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const distPath = path.join(__dirname, '..', 'dist');

app.set('trust proxy', 1);
app.use(express.json({ limit: '30mb' }));
app.use(
  cors({
    origin: frontendUrl,
    credentials: true
  })
);
if (process.env.NODE_ENV === 'production') {
  console.warn('[session] Usando MemoryStore em produção. Sessões Autodesk serão perdidas após restart/redeploy. Configure um session store persistente (Redis/Postgres).');
  if (!process.env.SESSION_STORE_PROVIDER) {
    console.warn('[session] Se o deploy reiniciar com frequência, os usuários precisarão fazer login novamente mesmo com cookie válido.');
  }
}

app.use(
  session({
    name: 'g5_instrumentos_session',
    secret: process.env.SESSION_SECRET || 'dev-only-change-this-session-secret',
    resave: false,
    saveUninitialized: false,
    rolling: true,
    cookie: {
      httpOnly: true,
      sameSite: 'lax',
      secure: 'auto',
      maxAge: sessionMaxAgeMs
    }
  })
);

function getApsCredentials() {
  const clientId = process.env.APS_CLIENT_ID;
  const clientSecret = process.env.APS_CLIENT_SECRET;

  // O segredo do app APS precisa existir somente no backend.
  if (!clientId || !clientSecret || !callbackUrl) {
    throw new Error(
      'Configuração APS incompleta. Confira APS_CLIENT_ID, APS_CLIENT_SECRET e APS_CALLBACK_URL no arquivo .env.'
    );
  }

  return { clientId, clientSecret };
}

function getBasicAuthorizationHeader(clientId, clientSecret) {
  // A Autodesk OAuth v2 espera Client ID e Client Secret no header Basic.
  const encodedCredentials = Buffer.from(`${clientId}:${clientSecret}`).toString('base64');
  return `Basic ${encodedCredentials}`;
}

async function exchangeCodeForToken(code) {
  const { clientId, clientSecret } = getApsCredentials();
  const body = new URLSearchParams({
    grant_type: 'authorization_code',
    code,
    redirect_uri: callbackUrl
  });

  const response = await fetch(apsTokenUrl, {
    method: 'POST',
    headers: {
      Authorization: getBasicAuthorizationHeader(clientId, clientSecret),
      'Content-Type': 'application/x-www-form-urlencoded',
      Accept: 'application/json'
    },
    body
  });

  if (!response.ok) {
    const details = await response.text();
    throw new Error(`A Autodesk recusou a troca do código por token. ${details}`);
  }

  return response.json();
}

async function refreshAccessToken(refreshToken) {
  const { clientId, clientSecret } = getApsCredentials();
  const body = new URLSearchParams({
    grant_type: 'refresh_token',
    refresh_token: refreshToken,
    scope: apsScopes.join(' ')
  });

  const response = await fetch(apsTokenUrl, {
    method: 'POST',
    headers: {
      Authorization: getBasicAuthorizationHeader(clientId, clientSecret),
      'Content-Type': 'application/x-www-form-urlencoded',
      Accept: 'application/json'
    },
    body
  });

  if (!response.ok) {
    throw new Error('Não foi possível renovar sua sessão Autodesk. Entre novamente.');
  }

  return response.json();
}

async function getValidAccessToken(req) {
  const token = req.session.apsToken;

  if (!token?.access_token) {
    return null;
  }

  const expiresSoon = token.expires_at && Date.now() > token.expires_at - 60_000;
  if (!expiresSoon) {
    return token.access_token;
  }

  if (!token.refresh_token) {
    return null;
  }

  // Se a sessão ainda tiver refresh_token, renovamos o access_token sem pedir novo login.
  const refreshedToken = await refreshAccessToken(token.refresh_token);
  req.session.apsToken = normalizeToken({
    ...refreshedToken,
    refresh_token: refreshedToken.refresh_token || token.refresh_token
  });
  return req.session.apsToken.access_token;
}

function normalizeToken(token) {
  return {
    ...token,
    expires_at: Date.now() + Number(token.expires_in || 0) * 1000
  };
}

function getSessionTokenScopes(req) {
  const scopeValue = req?.session?.apsToken?.scope || req?.session?.apsToken?.scopes || '';
  if (Array.isArray(scopeValue)) return scopeValue;
  return String(scopeValue || '')
    .split(/\s+/)
    .map((scope) => scope.trim())
    .filter(Boolean);
}

function hasWriteTokenScopes(req) {
  const scopes = new Set(getSessionTokenScopes(req));
  // Alguns retornos OAuth nao devolvem o campo scope. Nesse caso nao bloqueamos preventivamente,
  // mas o erro bruto da Autodesk sera mostrado no salvamento.
  if (!scopes.size) return true;
  return scopes.has('data:create') && scopes.has('data:write');
}

function formatTokenScopeHint(req) {
  const scopes = getSessionTokenScopes(req);
  return scopes.length ? ` Escopos atuais do token: ${scopes.join(' ')}.` : ' O token atual nao informou os escopos retornados pela Autodesk.';
}

function saveSession(req) {
  return new Promise((resolve, reject) => {
    req.session.save((error) => {
      if (error) {
        reject(error);
        return;
      }

      resolve();
    });
  });
}


function readNameFromAccessToken(accessToken) {
  try {
    const [, payload] = accessToken.split('.');
    if (!payload) {
      return null;
    }

    const decodedPayload = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));
    return decodedPayload.name || decodedPayload.given_name || decodedPayload.email || null;
  } catch {
    return null;
  }
}

function readIdentityFromAccessToken(accessToken) {
  try {
    const [, payload] = accessToken.split('.');
    if (!payload) {
      return {};
    }

    const decodedPayload = JSON.parse(Buffer.from(payload, 'base64url').toString('utf8'));

    return {
      name: decodedPayload.name || decodedPayload.given_name || decodedPayload.email || null,
      email: decodedPayload.email || null,
      ids: [
        decodedPayload.sub,
        decodedPayload.oxygen_id,
        decodedPayload.oxygenId,
        decodedPayload.user_id,
        decodedPayload.userId,
        decodedPayload.uid
      ].filter(Boolean)
    };
  } catch {
    return {};
  }
}


const APS_RETRYABLE_STATUS_CODES = new Set([502, 503, 504]);

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function fetchWithRetry(url, options = {}, retryOptions = {}) {
  const maxRetries = retryOptions.maxRetries ?? 2;
  const initialDelayMs = retryOptions.initialDelayMs ?? 600;

  for (let attempt = 0; attempt <= maxRetries; attempt += 1) {
    const response = await fetch(url, options);

    if (!APS_RETRYABLE_STATUS_CODES.has(response.status) || attempt === maxRetries) {
      return response;
    }

    await wait(initialDelayMs * (attempt + 1));
  }

  return fetch(url, options);
}

async function callApsApi(req, endpoint, options = {}) {
  const accessToken = await getValidAccessToken(req);

  if (!accessToken) {
    const error = new Error('Você precisa entrar com sua conta Autodesk primeiro. Sua sessão pode ter expirado; clique em "Entrar com Autodesk" novamente.');
    error.status = 401;
    throw error;
  }

  const response = await fetchWithRetry(`${apsApiBaseUrl}${endpoint}`, {
    method: options.method || 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      ...(options.body ? { 'Content-Type': 'application/json' } : {})
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });

  if (!response.ok) {
    const details = await response.text();
    const error = new Error(`Erro ao consultar a Autodesk em ${endpoint}. HTTP ${response.status}. ${formatApsError(details)}`);
    error.status = response.status;
    error.endpoint = endpoint;
    error.details = details;
    throw error;
  }

  return response.json();
}

async function callApsUrl(req, url, options = {}) {
  const accessToken = await getValidAccessToken(req);

  if (!accessToken) {
    const error = new Error('Você precisa entrar com sua conta Autodesk primeiro. Sua sessão pode ter expirado; clique em "Entrar com Autodesk" novamente.');
    error.status = 401;
    throw error;
  }

  const response = await fetchWithRetry(url, {
    method: options.method || 'GET',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      Accept: 'application/json',
      ...(options.body ? { 'Content-Type': 'application/json' } : {})
    },
    body: options.body ? JSON.stringify(options.body) : undefined
  });

  if (!response.ok) {
    const details = await response.text();
    const error = new Error(`Erro ao consultar a Autodesk em ${url}. HTTP ${response.status}. ${formatApsError(details)}`);
    error.status = response.status;
    error.endpoint = url;
    error.details = details;
    throw error;
  }

  return response.json();
}

async function fetchAllIssuePages(req, projectId, limit = 200) {
  const safeLimit = Math.min(Math.max(Number(limit) || 200, 1), 200);
  let endpoint = `/construction/issues/v1/projects/${projectId}/issues?limit=${safeLimit}`;
  const allIssues = [];
  const visited = new Set();
  let pages = 0;

  while (endpoint && !visited.has(endpoint)) {
    visited.add(endpoint);
    const result = await callApsApi(req, endpoint);
    const pageItems = result.results || result.data || [];
    allIssues.push(...pageItems);
    pages += 1;

    const nextLink =
      result.pagination?.nextUrl ||
      result.pagination?.next ||
      result.links?.next?.href ||
      result._links?.next?.href ||
      null;
    if (!nextLink) break;
    endpoint = nextLink.startsWith('http') ? nextLink.replace(apsApiBaseUrl, '') : nextLink;
  }

  return { issues: allIssues, pages };
}

function formatApsError(details) {
  const detailsText = String(details || '');

  if (detailsText.trim().toLowerCase().startsWith('<!doctype html') || detailsText.trim().startsWith('<')) {
    const titleMatch = detailsText.match(/<title>(.*?)<\/title>/i);
    const title = titleMatch?.[1] ? ` (${titleMatch[1]})` : '';
    return `A Autodesk ou o servidor intermediario retornou uma pagina de erro${title}. Tente atualizar novamente em alguns segundos.`;
  }

  try {
    const parsedDetails = JSON.parse(detailsText);
    const firstDetail = parsedDetails.details?.[0];
    const message = firstDetail?.message || parsedDetails.developerMessage || parsedDetails.message || detailsText;

    if (String(message).includes('assignedToType')) {
      return 'A Autodesk exigiu o tipo do responsavel. Ajustei o app para enviar responsavel como usuario.';
    }

    if (String(message).includes('must match format') && String(details).includes('dueDate')) {
      return 'A Autodesk recusou o formato do prazo. Use uma data valida no campo Prazo.';
    }

    return message;
  } catch {
    return detailsText;
  }
}


function isApsQuotaLimitError(error) {
  const message = `${error?.message || ''} ${error?.status || ''}`.toLowerCase();
  return error?.status === 429 || message.includes('quota') || message.includes('too many requests') || message.includes('rate limit');
}

function getAccProjectId(projectId) {
  // A Data Management API retorna projetos ACC com prefixo "b."; a Issues API usa o GUID puro.
  return projectId.startsWith('b.') ? projectId.slice(2) : projectId;
}

function readIssueField(issue, fieldNames, fallback = null) {
  for (const fieldName of fieldNames) {
    const value = issue[fieldName] ?? issue.attributes?.[fieldName];
    if (value !== undefined && value !== null && value !== '') {
      return value;
    }
  }

  return fallback;
}

function readNestedName(value) {
  if (!value) {
    return null;
  }

  if (typeof value === 'string') {
    return value;
  }

  return value.name || value.title || value.displayName || value.id || null;
}

function mapProjectUsers(result) {
  const users = new Map();

  for (const user of result.results || result.data || []) {
    const name = user.name || user.displayName || [user.firstName, user.lastName].filter(Boolean).join(' ') || user.email;
    const ids = [
      user.id,
      user.userId,
      user.autodeskId,
      user.autodesk_id,
      user.uid,
      user.email
    ].filter(Boolean);

    for (const id of ids) {
      users.set(id, name || id);
    }
  }

  return users;
}

function listProjectUsers(result) {
  return (result.results || result.data || [])
    .map((user) => {
      // Para atribuir uma issue, a Autodesk espera o Autodesk ID do usuario, nao o ID interno do membro no projeto.
      const id = user.autodeskId || user.autodesk_id || user.uid || user.id || user.userId || user.email;
      const name = user.name || user.displayName || [user.firstName, user.lastName].filter(Boolean).join(' ') || user.email || id;
      const role =
        user.role ||
        user.projectRole ||
        user.project_role ||
        user.roleName ||
        user.role_name ||
        user.accessLevel ||
        user.access_level ||
        user.permissionLevel ||
        user.permission_level ||
        '';
      const roleDetails = [
        user.roles,
        user.projectRoles,
        user.project_roles,
        user.products,
        user.services,
        user.access,
        user.permissions
      ]
        .filter(Boolean)
        .map((value) => {
          try {
            return JSON.stringify(value);
          } catch {
            return String(value);
          }
        })
        .join(' ');
      const projectAccessText = [role, roleDetails].filter(Boolean).join(' ');

      return {
        id,
        name,
        email: user.email || null,
        role: projectAccessText,
        accessLevel: user.accessLevel || user.access_level || user.permissionLevel || user.permission_level || '',
        isProjectAdmin: Boolean(
          user.isProjectAdmin ||
          user.projectAdmin ||
          user.isAdmin ||
          user.admin ||
          String(projectAccessText).toLowerCase().includes('admin') ||
          String(projectAccessText).toLowerCase().includes('administrador')
        )
      };
    })
    .filter((user) => user.id && user.name)
    .sort((firstUser, secondUser) => firstUser.name.localeCompare(secondUser.name, 'pt-BR', { sensitivity: 'base' }));
}

function mapAssigneeName(value, projectUsers = new Map()) {
  if (!value) {
    return null;
  }

  if (Array.isArray(value)) {
    return value.map((item) => mapAssigneeName(item, projectUsers)).filter(Boolean).join(', ') || null;
  }

  if (typeof value === 'string') {
    return projectUsers.get(value) || value;
  }

  const directName = value.name || value.displayName || [value.firstName, value.lastName].filter(Boolean).join(' ') || value.email;
  if (directName) {
    return directName;
  }

  const id = value.id || value.userId || value.autodeskId || value.autodesk_id || value.uid;
  return projectUsers.get(id) || id || null;
}

function normalizeIssueUserList(value, projectUsers = new Map()) {
  const list = Array.isArray(value) ? value : value ? [value] : [];
  const uniqueUsers = new Map();

  for (const userRef of list) {
    if (!userRef) continue;
    const id = typeof userRef === 'object'
      ? userRef.id || userRef.userId || userRef.autodeskId || userRef.autodesk_id || userRef.uid || userRef.email
      : userRef;
    const name = mapAssigneeName(userRef, projectUsers);
    const email = typeof userRef === 'object' ? userRef.email || userRef.mail || '' : '';
    const key = String(id || email || name || '').trim();
    if (!key) continue;
    uniqueUsers.set(key, {
      id: id || key,
      name: name || key,
      email
    });
  }

  return Array.from(uniqueUsers.values());
}

function mapUserName(value, projectUsers = new Map()) {
  if (!value) {
    return null;
  }

  if (typeof value === 'string') {
    return projectUsers.get(value) || value;
  }

  return (
    value.name ||
    value.displayName ||
    [value.firstName, value.lastName].filter(Boolean).join(' ') ||
    value.email ||
    projectUsers.get(value.id || value.userId || value.autodeskId || value.autodesk_id || value.uid) ||
    value.id ||
    value.userId ||
    null
  );
}

function getDefinitionOptions(definition) {
  const options =
    definition?.metadata?.list?.options ||
    definition?.metadata?.options ||
    definition?.options ||
    definition?.permittedValues ||
    [];

  return options
    .filter((item) => isActiveConfiguration(item))
    .map((item) => ({
      id: item.id || item.valueId || item.key || item.value,
      label: item.value || item.label || item.name || item.id
    }));
}

function isActiveConfiguration(item) {
  const status = String(item?.status || item?.state || item?.attributes?.status || '').toLowerCase();
  const deleted = item?.deleted || item?.isDeleted || item?.archived || item?.isArchived;
  const inactive = item?.inactive || item?.isInactive || item?.disabled || item?.isDisabled || item?.isActive === false;

  return !deleted && !inactive && !['inactive', 'disabled', 'archived', 'deleted'].includes(status);
}

function isActiveProject(project) {
  const attributes = project.attributes || {};
  const extensionData = attributes.extension?.data || {};
  const status = String(
    attributes.status ||
      attributes.projectStatus ||
      extensionData.status ||
      extensionData.projectStatus ||
      extensionData.state ||
      ''
  ).toLowerCase();

  const archived = attributes.archived || attributes.isArchived || extensionData.archived || extensionData.isArchived;
  const deleted = attributes.deleted || attributes.isDeleted || extensionData.deleted || extensionData.isDeleted;

  return !archived && !deleted && !['archived', 'inactive', 'deleted', 'suspended'].includes(status);
}

function isTemplateProject(project) {
  const attributes = project.attributes || {};
  const extension = attributes.extension || {};
  const extensionData = extension.data || {};
  const name = String(attributes.name || attributes.displayName || project.id || '').toLowerCase();
  const typeValues = [
    project.type,
    extension.type,
    extensionData.type,
    extensionData.projectType,
    extensionData.classification
  ]
    .filter(Boolean)
    .join(' ')
    .toLowerCase();

  return (
    typeValues.includes('template') ||
    extensionData.isTemplate === true ||
    attributes.isTemplate === true ||
    /^template\b/i.test(name) ||
    /\btemplate\b/i.test(name)
  );
}

function normalizeBusinessUnitLabel(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const normalized = text
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();

  if (normalized.includes('instrument')) return 'G5 Instrumentos';
  if (normalized.includes('engenharia')) return 'G5 Engenharia';

  return text;
}

function findProjectBusinessUnitValue(value, depth = 0) {
  if (!value || depth > 4) return '';

  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findProjectBusinessUnitValue(item, depth + 1);
      if (found) return found;
    }
    return '';
  }

  if (typeof value !== 'object') return '';

  for (const [key, rawValue] of Object.entries(value)) {
    const normalizedKey = String(key || '')
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '')
      .toLowerCase()
      .replace(/[^a-z0-9]/g, '');

    const isBusinessUnitKey =
      normalizedKey.includes('unidadedenegocios') ||
      normalizedKey.includes('businessunit') ||
      normalizedKey.includes('negocio');

    if (isBusinessUnitKey) {
      const candidate =
        typeof rawValue === 'object'
          ? rawValue?.value || rawValue?.displayValue || rawValue?.label || rawValue?.name || rawValue?.text
          : rawValue;
      const normalized = normalizeBusinessUnitLabel(candidate);
      if (normalized) return normalized;
    }

    const nested = findProjectBusinessUnitValue(rawValue, depth + 1);
    if (nested) return nested;
  }

  return '';
}

function getProjectBusinessUnitId(project) {
  const attributes = project.attributes || {};
  const extensionData = attributes.extension?.data || {};
  const adminData = project.adminData || {};
  const candidates = [
    project.businessUnitId,
    project.business_unit_id,
    project.unidadeNegociosId,
    adminData.businessUnitId,
    adminData.business_unit_id,
    adminData.unidadeNegociosId,
    adminData.relationships?.businessUnit?.data?.id,
    adminData.relationships?.business_unit?.data?.id,
    attributes.businessUnitId,
    attributes.business_unit_id,
    attributes.unidadeNegociosId,
    attributes.relationships?.businessUnit?.data?.id,
    attributes.relationships?.business_unit?.data?.id,
    extensionData.businessUnitId,
    extensionData.business_unit_id,
    extensionData.unidadeNegociosId,
    extensionData.relationships?.businessUnit?.data?.id,
    extensionData.relationships?.business_unit?.data?.id
  ];

  return candidates.find((candidate) => candidate !== undefined && candidate !== null && String(candidate).trim() !== '') || '';
}

function getProjectBusinessUnit(project, businessUnitNamesById = new Map()) {
  const attributes = project.attributes || {};
  const extensionData = attributes.extension?.data || {};
  const adminData = project.adminData || {};
  const businessUnitId = normalizeAccProjectId(getProjectBusinessUnitId(project));
  const candidates = [
    project.businessUnit,
    project.businessUnitName,
    project.business_unit,
    project.unidadeNegocios,
    project.unidade_de_negocios,
    project['Unidade de negocios'],
    adminData.businessUnit,
    adminData.businessUnitName,
    adminData.business_unit,
    adminData.unidadeNegocios,
    adminData.unidade_de_negocios,
    adminData['Unidade de negocios'],
    attributes.businessUnit,
    attributes.businessUnitName,
    attributes.business_unit,
    attributes.unidadeNegocios,
    attributes.unidade_de_negocios,
    attributes['Unidade de negocios'],
    attributes['Unidade de negócios'],
    extensionData.businessUnit,
    extensionData.businessUnitName,
    extensionData.business_unit,
    extensionData.unidadeNegocios,
    extensionData.unidade_de_negocios,
    extensionData['Unidade de negocios'],
    extensionData['Unidade de negócios'],
    businessUnitNamesById.get(businessUnitId),
    findProjectBusinessUnitValue(adminData),
    findProjectBusinessUnitValue(attributes.customAttributes),
    findProjectBusinessUnitValue(extensionData),
    findProjectBusinessUnitValue(project)
  ];

  return candidates.map(normalizeBusinessUnitLabel).find(Boolean) || 'Nao classificado';
}

function normalizeAccProjectId(value) {
  return String(value || '').replace(/^b\./i, '').trim().toLowerCase();
}

function getAccountIdFromHubId(hubId) {
  return normalizeAccProjectId(hubId);
}

function collectProjectRecords(payload) {
  if (!payload) return [];
  if (Array.isArray(payload)) return payload;

  const containers = [
    payload.data,
    payload.results,
    payload.items,
    payload.projects,
    payload.data?.results,
    payload.data?.items,
    payload.data?.projects
  ];

  for (const container of containers) {
    if (Array.isArray(container)) return container;
  }

  return [];
}

function registerAdminProjectMetadata(map, record) {
  if (!record || typeof record !== 'object') return;

  const idCandidates = [
    record.id,
    record.projectId,
    record.project_id,
    record.projectGuid,
    record.projectGuidId,
    record.attributes?.id,
    record.attributes?.projectId,
    record.attributes?.projectGuid,
    record.attributes?.extension?.data?.projectId,
    record.attributes?.extension?.data?.projectGuid
  ].filter(Boolean);

  for (const id of idCandidates) {
    const normalized = normalizeAccProjectId(id);
    if (normalized) map.set(normalized, record);
  }
}

function registerBusinessUnitName(map, record) {
  if (!record || typeof record !== 'object') return;

  const id = normalizeAccProjectId(record.id || record.businessUnitId || record.business_unit_id || record.attributes?.id);
  const name =
    record.name ||
    record.displayName ||
    record.businessUnitName ||
    record.attributes?.name ||
    record.attributes?.displayName ||
    record.attributes?.businessUnitName ||
    record.attributes?.extension?.data?.name ||
    record.attributes?.extension?.data?.businessUnitName ||
    '';

  const normalizedName = normalizeBusinessUnitLabel(name);
  if (id && normalizedName) map.set(id, normalizedName);
}

async function getBusinessUnitNamesById(req, accountId) {
  const businessUnitNamesById = new Map();
  if (!accountId) return businessUnitNamesById;

  const endpoints = [
    `/construction/admin/v1/accounts/${encodeURIComponent(accountId)}/business-units?limit=200`,
    `/construction/admin/v1/accounts/${encodeURIComponent(accountId)}/businessUnits?limit=200`
  ];

  for (const endpoint of endpoints) {
    try {
      const result = await callApsApi(req, endpoint);
      for (const record of collectProjectRecords(result)) {
        registerBusinessUnitName(businessUnitNamesById, record);
      }
      if (businessUnitNamesById.size) break;
    } catch (error) {
      console.warn('Nao foi possivel ler unidades de negocio pelo ACC Admin API.', error.message);
    }
  }

  return businessUnitNamesById;
}

async function getAdminProjectMetadataMap(req, hubId, projects = []) {
  const accountId = getAccountIdFromHubId(hubId);
  const metadataByProjectId = new Map();
  if (!accountId) return metadataByProjectId;

  try {
    const listResult = await callApsApi(
      req,
      `/construction/admin/v1/accounts/${encodeURIComponent(accountId)}/projects?limit=200`
    );
    for (const record of collectProjectRecords(listResult)) {
      registerAdminProjectMetadata(metadataByProjectId, record);
    }
  } catch (error) {
    console.warn('Nao foi possivel ler projetos pelo ACC Admin API.', error.message);
  }

  const projectsMissingBusinessUnit = projects.filter((project) => {
    const normalizedProjectId = normalizeAccProjectId(project.id);
    const adminData = metadataByProjectId.get(normalizedProjectId);
    return !adminData || getProjectBusinessUnit({ ...project, adminData }) === 'Nao classificado';
  });

  const detailResults = await Promise.all(
    projectsMissingBusinessUnit.map(async (project) => {
      const normalizedProjectId = normalizeAccProjectId(project.id);
      if (!normalizedProjectId) return null;

      try {
        const detail = await callApsApi(
          req,
          `/construction/admin/v1/accounts/${encodeURIComponent(accountId)}/projects/${encodeURIComponent(normalizedProjectId)}`
        );
        return { normalizedProjectId, detail };
      } catch {
        return null;
      }
    })
  );

  for (const result of detailResults) {
    if (result?.detail) {
      metadataByProjectId.set(result.normalizedProjectId, result.detail);
      registerAdminProjectMetadata(metadataByProjectId, result.detail);
    }
  }

  return metadataByProjectId;
}

async function isCurrentUserProjectMember(req, projectId) {
  const accessToken = await getValidAccessToken(req);
  const identity = readIdentityFromAccessToken(accessToken || '');
  const identityIds = new Set((identity.ids || []).map(String));
  const identityEmail = String(identity.email || '').toLowerCase();

  if (!identityIds.size && !identityEmail) {
    return true;
  }

  const accProjectId = getAccProjectId(projectId);
  const result = await callApsApi(req, `/construction/admin/v1/projects/${accProjectId}/users?limit=200`);

  return (result.results || result.data || []).some((user) => {
    const userEmail = String(user.email || '').toLowerCase();
    const userIds = [
      user.id,
      user.userId,
      user.autodeskId,
      user.autodesk_id,
      user.uid
    ]
      .filter(Boolean)
      .map(String);

    return (identityEmail && userEmail === identityEmail) || userIds.some((id) => identityIds.has(id));
  });
}

function normalizeCustomAttributes(rawCustomAttributes, attributeDefinitions = new Map()) {
  if (!rawCustomAttributes) {
    return [];
  }

  const findAttributeDefinition = (field = {}, fallbackKey = '') => {
    const keys = [
      field.attributeDefinitionId,
      field.definitionId,
      field.fieldId,
      field.id,
      field.key,
      field.name,
      field.title,
      field.displayName,
      field.label,
      fallbackKey
    ].filter((key) => key !== undefined && key !== null && key !== '');
    for (const key of keys) {
      const definition = attributeDefinitions.get(String(key));
      if (definition) return definition;
    }
    return null;
  };

  if (Array.isArray(rawCustomAttributes)) {
    return rawCustomAttributes
      .map((field) => {
        const id = field.attributeDefinitionId || field.definitionId || field.fieldId || field.id || field.key || field.name;
        const definition = findAttributeDefinition(field, id);
        const rawValue = field.value ?? field.displayValue ?? field.text ?? field.selectedValue ?? null;
        const readableValue = getDefinitionOptionLabelSafe(definition, rawValue) || rawValue;

        return {
          id,
          name: field.name || field.title || field.displayName || definition?.title || definition?.name || id || 'Campo personalizado',
          value: readableValue,
          rawValue,
          type: definition?.dataType || definition?.type || field.type || null,
          options: getDefinitionOptions(definition)
        };
      })
      .filter((field) => field.value !== null && field.value !== undefined && field.value !== '');
  }

  if (typeof rawCustomAttributes === 'object') {
    return Object.entries(rawCustomAttributes)
      .map(([key, value]) => {
        const definition = findAttributeDefinition(value && typeof value === 'object' ? value : {}, key);
        const rawValue =
          typeof value === 'object' && value !== null ? value.value ?? value.displayValue ?? value.selectedValue ?? value.id : value;
        const readableValue = getDefinitionOptionLabelSafe(definition, rawValue) || rawValue;

        return {
          id: key,
          name: definition?.title || definition?.name || key,
          value: typeof readableValue === 'object' && readableValue !== null ? JSON.stringify(readableValue) : readableValue,
          rawValue,
          type: definition?.dataType || definition?.type || null,
          options: getDefinitionOptions(definition)
        };
      })
      .filter((field) => field.value !== null && field.value !== undefined && field.value !== '');
  }

  return [];
}

function mapIssue(
  issue,
  attributeDefinitions = new Map(),
  issueTypeMaps = { types: new Map(), subtypes: new Map() },
  projectUsers = new Map()
) {
  const issueTypeId = readIssueField(issue, ['issueTypeId', 'typeId']);
  const issueSubtypeId = readIssueField(issue, ['issueSubtypeId', 'subtypeId']);
  const issueType = readIssueField(issue, ['issueType', 'issueTypeName', 'type']);
  const issueSubtype = readIssueField(issue, ['issueSubtype', 'issueSubtypeName', 'subtype']);
  const assignedTo = readIssueField(issue, ['assignedTo', 'assignedToId', 'assignee', 'assigneeName']);
  const watchers = readIssueField(issue, ['watchers', 'watcherIds', 'followers', 'followerIds'], []);
  const category = readIssueField(issue, ['category', 'categoryName', 'rootCauseCategory']);
  const customAttributes = readIssueField(issue, ['customAttributes', 'custom_attributes', 'customFields', 'custom_fields'], []);
  const typeDefinition = issueTypeMaps.types.get(issueTypeId);
  const subtypeDefinition = issueTypeMaps.subtypes.get(issueSubtypeId);

  return {
    id: issue.id,
    displayId: readIssueField(issue, ['displayId', 'identifier', 'number']),
    title: readIssueField(issue, ['title'], 'Issue sem titulo'),
    description: readIssueField(issue, ['description', 'details'], ''),
    status: readIssueField(issue, ['status'], 'unknown'),
    issueTypeId,
    issueSubtypeId,
    category: readNestedName(category) || typeDefinition?.category || typeDefinition?.title || null,
    issueType: readNestedName(issueSubtype) || subtypeDefinition?.title || readNestedName(issueType) || typeDefinition?.title || null,
    issueSubtype: null,
    assignedTo: mapAssigneeName(assignedTo, projectUsers),
    watchers: normalizeIssueUserList(watchers, projectUsers),
    followers: normalizeIssueUserList(watchers, projectUsers),
    dueDate: readIssueField(issue, ['dueDate', 'due_date']),
    startDate: readIssueField(issue, ['startDate', 'start_date']),
    openedAt: readIssueField(issue, ['openedAt', 'opened_at']),
    closedAt: readIssueField(issue, ['closedAt', 'closed_at']),
    createdAt: readIssueField(issue, ['createdAt', 'created_at', 'createdDate']),
    updatedAt: readIssueField(issue, ['updatedAt', 'updated_at', 'updatedDate']),
    createdBy: mapUserName(readIssueField(issue, ['createdBy', 'created_by']), projectUsers),
    updatedBy: mapUserName(readIssueField(issue, ['updatedBy', 'updated_by']), projectUsers),
    openedBy: mapUserName(readIssueField(issue, ['openedBy', 'opened_by']), projectUsers),
    closedBy: mapUserName(readIssueField(issue, ['closedBy', 'closed_by']), projectUsers),
    location: readIssueField(issue, ['locationDetails', 'location', 'locationId']),
    owner: mapUserName(readIssueField(issue, ['ownerId', 'owner']), projectUsers),
    published: readIssueField(issue, ['published']),
    commentCount: readIssueField(issue, ['commentCount'], 0),
    attachmentCount: readIssueField(issue, ['attachmentCount'], 0),
    customAttributes: normalizeCustomAttributes(customAttributes, attributeDefinitions),
    raw: issue
  };
}

function mapIssueComment(comment, projectUsers = new Map()) {
  const body = readIssueField(comment, ['body', 'text', 'comment', 'message', 'description'], '');
  const createdBy = readIssueField(comment, ['createdBy', 'created_by', 'author', 'user']);

  return {
    id: comment.id || comment.commentId || crypto.randomUUID(),
    body: typeof body === 'object' && body !== null ? body.value || body.text || JSON.stringify(body) : body,
    createdAt: readIssueField(comment, ['createdAt', 'created_at', 'createdDate', 'updatedAt', 'updated_at']),
    createdBy: mapUserName(createdBy, projectUsers),
    updatedAt: readIssueField(comment, ['updatedAt', 'updated_at']),
    raw: comment
  };
}

async function getIssueComments(req, projectId, issueId, projectUsers) {
  const encodedIssueId = encodeURIComponent(issueId);
  const attempts = [
    `/construction/issues/v1/projects/${projectId}/issues/${encodedIssueId}/comments?limit=200`,
    `/construction/issues/v1/projects/${projectId}/comments?filter[issueId]=${encodedIssueId}&limit=200`
  ];

  let lastError = null;

  for (const endpoint of attempts) {
    try {
      const result = await callApsApi(req, endpoint);
      return {
        comments: (result.results || result.data || result.comments || []).map((comment) =>
          mapIssueComment(comment, projectUsers)
        ),
        warning: ''
      };
    } catch (error) {
      lastError = error;
    }
  }

  return {
    comments: [],
    warning: lastError?.message || 'Nao foi possivel consultar os comentarios desta issue.'
  };
}

function mapAttributeDefinitions(result) {
  const definitions = new Map();
  const registerDefinition = (key, definition) => {
    if (key === undefined || key === null || key === '') return;
    definitions.set(String(key), definition);
  };

  for (const definition of result.results || result.data || []) {
    [
      definition.id,
      definition.attributeDefinitionId,
      definition.definitionId,
      definition.key,
      definition.name,
      definition.title,
      definition.displayName,
      definition.label
    ].forEach((key) => registerDefinition(key, definition));
  }

  return definitions;
}

function mapIssueTypes(result) {
  const maps = { types: new Map(), subtypes: new Map() };

  for (const type of result.results || result.data || []) {
    const typeId = type.id || type.issueTypeId;
    const typeTitle = type.title || type.name || type.displayName;
    if (typeId) {
      maps.types.set(typeId, {
        id: typeId,
        title: typeTitle,
        category: type.category || type.categoryName || typeTitle
      });
    }

    for (const subtype of type.subtypes || type.issueSubtypes || []) {
      const subtypeId = subtype.id || subtype.issueSubtypeId;
      if (subtypeId) {
        maps.subtypes.set(subtypeId, {
          id: subtypeId,
          title: subtype.title || subtype.name || subtype.displayName,
          typeId,
          category: type.category || type.categoryName || typeTitle
        });
      }
    }
  }

  return maps;
}

function listIssueTypes(result) {
  const items = [];
  const categoryNames = new Set();

  for (const type of (result.results || result.data || []).filter(isActiveConfiguration)) {
    const typeId = type.id || type.issueTypeId;
    const typeTitle = type.title || type.name || type.displayName;
    const categoryName = type.category || type.categoryName || typeTitle;

    if (categoryName) categoryNames.add(categoryName);

    items.push({
      id: typeId,
      title: typeTitle,
      category: categoryName,
      kind: 'type'
    });

    for (const subtype of (type.subtypes || type.issueSubtypes || []).filter(isActiveConfiguration)) {
      items.push({
        id: subtype.id || subtype.issueSubtypeId,
        title: subtype.title || subtype.name || subtype.displayName,
        category: categoryName,
        typeId,
        typeTitle,
        kind: 'subtype'
      });
    }
  }

  for (const categoryName of categoryNames) {
    items.push({
      id: `category::${normalizeComparableName(categoryName)}`,
      title: categoryName,
      category: categoryName,
      kind: 'category'
    });
  }

  return items.filter((item) => item.id && item.title);
}

function findIssueTypeByName(issueTypes, requestedName) {
  const normalizedRequestedName = String(requestedName || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

  return issueTypes.find((item) => {
    const values = [item.title, item.category, item.typeTitle].filter(Boolean);
    return values.some((value) =>
      String(value).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().includes(normalizedRequestedName)
    );
  });
}

const manualIssueCatalog = [
  { code: 'INF', type: 'Solicitação de Informação', category: 'Interface e Coordenação Multidisciplinar' },
  { code: 'INT', type: 'Interface Multidisciplinar', category: 'Interface e Coordenação Multidisciplinar' },
  { code: 'DTP', type: 'Definição Técnica Pendente', category: 'Interface e Coordenação Multidisciplinar' },
  { code: 'DEC', type: 'Decisão de Projeto', category: 'Interface e Coordenação Multidisciplinar' },
  { code: 'COM', type: 'Compatibilização', category: 'Interface e Coordenação Multidisciplinar' },
  { code: 'SGI', type: 'Solicitação da Gestão Interna', category: 'Gestão / Cliente Interno' },
  { code: 'EVI', type: 'Evidência / Documento de Apoio', category: 'Gestão / Cliente Interno' },
  { code: 'ENC', type: 'Encaminhamento Interno', category: 'Gestão / Cliente Interno' },
  { code: 'PRI', type: 'Priorização / Direcionamento', category: 'Gestão / Cliente Interno' },
  { code: 'CLI', type: 'Solicitação do Cliente', category: 'Cliente Externo' },
  { code: 'RCL', type: 'Resposta ao Cliente', category: 'Cliente Externo' },
  { code: 'PCL', type: 'Pendência do Cliente', category: 'Cliente Externo' },
  { code: 'DCL', type: 'Decisão do Cliente', category: 'Cliente Externo' },
  { code: 'TER', type: 'Terceiros / Fornecedores', category: 'Terceiros / Fornecedores' },
  { code: 'SOT', type: 'Solicitação para Terceiros', category: 'Terceiros / Fornecedores' },
  { code: 'RET', type: 'Retorno de Terceiros', category: 'Terceiros / Fornecedores' },
  { code: 'MOB', type: 'Mobilização em Campo', category: 'Mobilização e Campo' },
  { code: 'LIB', type: 'Liberação de Acesso', category: 'Mobilização e Campo' },
  { code: 'SEG', type: 'Segurança / Integração', category: 'Mobilização e Campo' },
  { code: 'CAM', type: 'Atividade de Campo', category: 'Mobilização e Campo' },
  { code: 'DCM', type: 'Documento de Campo', category: 'Mobilização e Campo' },
  { code: 'CCM', type: 'Pendência Cliente Campo', category: 'Mobilização e Campo' }
];

const manualIssueCatalogByCode = new Map(manualIssueCatalog.map((item) => [item.code, item]));

function normalizeSpreadsheetKey(value) {
  return normalizeDocumentText(value).replace(/\s+/g, ' ');
}

function makeNormalizedIssueKey(...values) {
  return normalizeSpreadsheetKey(values.filter(Boolean).join(' ')).replace(/\s+/g, '');
}

function findColumnValue(row, aliases) {
  const normalizedAliases = aliases.map(normalizeSpreadsheetKey);
  const exact = Object.entries(row).find(([key]) => normalizedAliases.includes(normalizeSpreadsheetKey(key)));
  if (exact) return exact[1];

  const partial = Object.entries(row).find(([key]) => {
    const normalizedKey = normalizeSpreadsheetKey(key);
    return normalizedAliases.some((alias) => normalizedKey.includes(alias) || alias.includes(normalizedKey));
  });

  return partial ? partial[1] : '';
}

function normalizeCellValue(value) {
  if (value instanceof Date) {
    return value.toISOString().slice(0, 10);
  }

  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function parseSpreadsheetDate(value) {
  if (!value) return '';
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toISOString().slice(0, 10);

  const text = String(value).trim();
  const numericDate = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
  if (numericDate) {
    const [, firstPart, secondPart, year] = numericDate;
    const firstNumber = Number(firstPart);
    const secondNumber = Number(secondPart);
    const fullYear = year.length === 2 ? `20${year}` : year;
    const day = secondNumber > 12 && firstNumber <= 12 ? secondPart : firstPart;
    const month = secondNumber > 12 && firstNumber <= 12 ? firstPart : secondPart;
    const parsed = new Date(Date.UTC(Number(fullYear), Number(month) - 1, Number(day)));

    if (!Number.isNaN(parsed.getTime()) && parsed.getUTCMonth() === Number(month) - 1) {
      return `${fullYear.padStart(4, '0')}-${month.padStart(2, '0')}-${day.padStart(2, '0')}`;
    }
  }

  const parsedDate = new Date(text);
  if (!Number.isNaN(parsedDate.getTime())) return parsedDate.toISOString().slice(0, 10);
  return '';
}

function getIssueCodeFromText(...values) {
  const text = values.filter(Boolean).join(' ');
  const match = text.match(/\b([A-Z]{2,4})[-\s]?\d+(?:[.\-]\d+)?\b/i);
  return match ? match[1].toUpperCase() : '';
}

function pickManualIssueRule(rowValues) {
  const explicitType = rowValues.issueType || rowValues.category || '';
  const code = getIssueCodeFromText(rowValues.issueCode, rowValues.title, rowValues.issueType, rowValues.description);
  const manualRule = manualIssueCatalogByCode.get(code);

  if (manualRule) return { ...manualRule, code };

  const normalizedType = normalizeSpreadsheetKey(explicitType);
  const byType = manualIssueCatalog.find(
    (item) => normalizeSpreadsheetKey(item.type).includes(normalizedType) || normalizedType.includes(normalizeSpreadsheetKey(item.type))
  );

  return byType ? { ...byType, code: byType.code } : null;
}

function buildIssueTitleFromEap(rowValues, manualRule) {
  if (rowValues.title) return rowValues.title;

  const typePrefix = manualRule?.code || manualRule?.type || rowValues.issueType || 'ISSUE';
  const code = rowValues.eapCode || rowValues.packageCode || 'SEM-CODIGO';
  const description = rowValues.description || rowValues.observations || 'Item da EAP';
  const compactDescription = description.length > 80 ? `${description.slice(0, 77)}...` : description;
  return `${typePrefix} | ${code} | ${compactDescription}`;
}

function mapImpactValue(value) {
  const normalized = normalizeSpreadsheetKey(value);
  if (!normalized) return '';
  if (normalized.includes('alto') || normalized.includes('critico')) return 'Alto';
  if (normalized.includes('medio') || normalized.includes('media')) return 'Médio';
  if (normalized.includes('baixo') || normalized.includes('baixa')) return 'Baixo';
  return value;
}

function fieldValueMatchesOption(option, value) {
  const normalizedValue = normalizeSpreadsheetKey(value);
  return [option.id, option.value, option.label, option.name]
    .filter(Boolean)
    .some((optionValue) => normalizeSpreadsheetKey(optionValue) === normalizedValue);
}

function buildCustomAttributesFromEapV2(rowValues, row, fieldDefinitions) {
  const reservedKeys = new Set(
    [
      'codigo eap',
      'item eap',
      'eap',
      'descricao',
      'descrição',
      'disciplina',
      'pacote',
      'awp',
      'cwp',
      'iwp',
      'tipo de issue',
      'tipo issue',
      'categoria',
      'nome do issue',
      'nome da issue',
      'status',
      'responsavel',
      'responsável',
      'data prevista g5',
      'prazo',
      'observacoes',
      'observações'
    ].map(normalizeSpreadsheetKey)
  );

  const fieldValues = {
    'Data Prevista G5': rowValues.dueDate,
    'Impacto em Prazo': mapImpactValue(rowValues.impactSchedule),
    'Impacto no Escopo': mapImpactValue(rowValues.impactScope),
    'Impacto em Medição': mapImpactValue(rowValues.impactMeasurement),
    'EAP vinculada': rowValues.eapCode,
    'Código AWP / WF': rowValues.packageCode,
    Disciplina: rowValues.discipline,
    Fase: rowValues.phase
  };

  for (const [key, value] of Object.entries(row)) {
    const cleanKey = String(key).replace(/__\d+$/, '').trim();
    const normalizedKey = normalizeSpreadsheetKey(cleanKey);
    if (!value || reservedKeys.has(normalizedKey)) continue;
    fieldValues[cleanKey] = normalizeCellValue(value);
  }

  return Object.entries(fieldValues)
    .map(([fieldName, rawValue]) => {
      const value = normalizeCellValue(rawValue);
      if (!value) return null;

      const definition = fieldDefinitions.find((field) => {
        const normalizedDefinition = normalizeSpreadsheetKey(field.name);
        const normalizedFieldName = normalizeSpreadsheetKey(fieldName);
        return normalizedDefinition === normalizedFieldName || normalizedDefinition.includes(normalizedFieldName) || normalizedFieldName.includes(normalizedDefinition);
      });

      if (!definition?.id) return null;

      const option = definition.options?.find((candidate) => fieldValueMatchesOption(candidate, value));
      return {
        attributeDefinitionId: definition.id,
        name: definition.name,
        value: option?.id || value,
        displayValue: option?.label || value
      };
    })
    .filter(Boolean);
}

function parseEapWorkbookV2(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = workbook.SheetNames.find((name) => normalizeSpreadsheetKey(name) === normalizeSpreadsheetKey('ACS Build'));

  if (!sheetName) {
    const error = new Error("A planilha enviada não possui a aba 'ACS Build'. Use a planilha padrão revisada para importação de issues.");
    error.status = 400;
    throw error;
  }

  const tableRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false });
  const knownAliases = [
    'atualizar?', 'atualizar', 'acao', 'ação',
    'issue id', 'id da issue',
    'title', 'titulo', 'título', 'nome do issue',
    'status', 'priority', 'due date', 'data prevista',
    'assigned to', 'responsavel', 'responsável',
    'description', 'descricao', 'descrição'
  ];
  const scoreHeaderRow = (row) =>
    row.reduce((score, cell) => {
      const normalized = normalizeSpreadsheetKey(cell);
      if (!normalized) return score;
      if (knownAliases.some((alias) => normalized === normalizeSpreadsheetKey(alias))) return score + 3;
      if (knownAliases.some((alias) => normalized.includes(normalizeSpreadsheetKey(alias)))) return score + 1;
      return score;
    }, 0);

  const headerCandidate = tableRows
    .map((row, index) => ({ index, score: scoreHeaderRow(row) }))
    .reduce((best, candidate) => (candidate.score > best.score ? candidate : best), { index: -1, score: 0 });
  const headerRowIndex = headerCandidate.score >= 3 ? headerCandidate.index : -1;

  if (headerRowIndex < 0) {
    const error = new Error('Encontrei a aba ACS Build, mas nao encontrei uma linha de cabecalho reconhecivel.');
    error.status = 400;
    throw error;
  }

  const headers = tableRows[headerRowIndex].map((header, index) => {
    const label = String(header || '').trim();
    return label ? `${label}__${index}` : `Coluna ${index + 1}`;
  });

  const rows = tableRows.slice(headerRowIndex + 1).map((row, rowIndex) => ({
    rowNumber: headerRowIndex + rowIndex + 2,
    values: headers.reduce((record, header, index) => {
      record[header] = row[index] ?? '';
      return record;
    }, {})
  }));

  return {
    sheetName,
    headerRowIndex,
    rows
  };
}

function readDocumentAttribute(version, names) {
  const attributes = version?.attributes || {};
  const extensionData = attributes.extension?.data || {};
  const customAttributes =
    extensionData.customAttributes ||
    extensionData.custom_attributes ||
    attributes.customAttributes ||
    attributes.custom_attributes ||
    {};

  for (const name of names) {
    const value = attributes[name] ?? extensionData[name] ?? customAttributes[name];
    if (value !== undefined && value !== null && value !== '') {
      return typeof value === 'object' ? value.value || value.displayValue || value.name || JSON.stringify(value) : value;
    }
  }

  const normalizedNames = names.map((name) => String(name).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase());

  for (const [key, value] of Object.entries(customAttributes)) {
    const normalizedKey = String(key).normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    if (normalizedNames.some((name) => normalizedKey.includes(name))) {
      return typeof value === 'object' ? value.value || value.displayValue || value.name || JSON.stringify(value) : value;
    }
  }

  return null;
}

function normalizeDocumentAttributeKey(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function stringifyDocumentAttributeValue(value) {
  if (value === undefined || value === null || value === '') return null;

  if (Array.isArray(value)) {
    const values = value.map(stringifyDocumentAttributeValue).filter(Boolean);
    return values.length ? values.join(', ') : null;
  }

  if (typeof value === 'object') {
    const readable =
      value.displayName ??
      value.display_name ??
      value.displayValue ??
      value.display_value ??
      value.display_value_string ??
      value.selectedValue ??
      value.selected_value ??
      value.selectedName ??
      value.selected_name ??
      value.optionName ??
      value.option_name ??
      value.optionValue ??
      value.option_value ??
      value.value ??
      value.val ??
      value.name ??
      value.label ??
      value.text ??
      value.title;
    return stringifyDocumentAttributeValue(readable);
  }

  return String(value).trim();
}

function collectDocumentAttributeCandidates(source, candidates = [], depth = 0, visited = new Set()) {
  if (!source || typeof source !== 'object' || depth > 5 || visited.has(source)) return candidates;

  visited.add(source);

  if (Array.isArray(source)) {
    for (const item of source) collectDocumentAttributeCandidates(item, candidates, depth + 1, visited);
    return candidates;
  }

  const labeledKey = source.name || source.displayName || source.display_name || source.title || source.label || source.key;
  const labeledValue =
    source.value ?? source.displayValue ?? source.display_value ?? source.selectedValue ?? source.selected_value ?? source.text;
  if (labeledKey && labeledValue !== undefined) {
    const key = normalizeDocumentAttributeKey(labeledKey);
    const value = stringifyDocumentAttributeValue(labeledValue);
    if (key && value) candidates.push({ key, compactKey: key.replace(/\s+/g, ''), value });
  }

  for (const [rawKey, rawValue] of Object.entries(source)) {
    const key = normalizeDocumentAttributeKey(rawKey);
    const value = stringifyDocumentAttributeValue(rawValue);
    if (key && value) candidates.push({ key, compactKey: key.replace(/\s+/g, ''), value });

    const shouldSearchInside = [
      'attributes',
      'custom attributes',
      'customattributes',
      'custom attributes values',
      'customattributesvalues',
      'custom attribute values',
      'customattributevalues',
      'custom',
      'fields',
      'field',
      'values',
      'details',
      'included',
      'relationships',
      'versions',
      'version',
      'properties',
      'metadata',
      'data',
      'extension'
    ].includes(key);

    const isSmallObject =
      rawValue &&
      typeof rawValue === 'object' &&
      !Array.isArray(rawValue) &&
      Object.keys(rawValue).length <= 40 &&
      !['links', 'href', 'url'].includes(key);

    if (rawValue && typeof rawValue === 'object' && (shouldSearchInside || (depth < 3 && isSmallObject))) {
      collectDocumentAttributeCandidates(rawValue, candidates, depth + 1, visited);
    }
  }

  return candidates;
}

function isInternalDocumentReviewValue(value) {
  const text = normalizeDocumentAttributeKey(value).replace(/\s+/g, '_');
  return [
    'not_in_review',
    'in_review',
    'review',
    'approved',
    'rejected',
    'pending',
    'closed',
    'open'
  ].includes(text);
}

function readDocumentAttributeFromSources(sources, names, options = {}) {
  const sourceList = Array.isArray(sources) ? sources : [sources];
  const candidates = [];

  for (const source of sourceList) {
    collectDocumentAttributeCandidates(source, candidates);
  }

  const normalizedNames = names.map(normalizeDocumentAttributeKey).filter(Boolean);
  const compactNames = normalizedNames.map((name) => name.replace(/\s+/g, ''));
  const excludedKeys = (options.excludeKeys || []).map(normalizeDocumentAttributeKey).filter(Boolean);
  const excludedCompactKeys = excludedKeys.map((name) => name.replace(/\s+/g, ''));

  for (const candidate of candidates) {
    const isExcluded =
      excludedKeys.some((name) => candidate.key === name || candidate.key.includes(name)) ||
      excludedCompactKeys.some((name) => candidate.compactKey === name || candidate.compactKey.includes(name));
    if (isExcluded) continue;

    const matches = options.exact
      ? normalizedNames.some((name) => candidate.key === name) || compactNames.some((name) => candidate.compactKey === name)
      : normalizedNames.some((name) => candidate.key === name || candidate.key.includes(name)) ||
        compactNames.some((name) => candidate.compactKey === name || candidate.compactKey.includes(name));

    if (matches) {
      if (options.rejectValue?.(candidate.value, candidate)) continue;
      return candidate.value;
    }
  }

  return null;
}

function inferDocumentRevisionFromName(value) {
  const text = String(value || '');
  const match = text.match(/(?:^|[-_\s])(?:REV|R)(?:\.|\s|-|_)?([A-Z0-9]{1,3})(?=$|[-_\s.])/i);
  return match ? match[1].toUpperCase() : '';
}

function inferDocumentGrdFromText(value) {
  const text = String(value || '');
  const match = text.match(/\bGRD[-_\s.]?([A-Z0-9-]{1,20})\b/i);
  return match ? `GRD-${match[1].replace(/^GRD[-_\s.]?/i, '').toUpperCase()}` : '';
}

function collectDocumentBatchRecords(result) {
  if (!result) return [];
  if (Array.isArray(result.results)) return result.results;
  if (Array.isArray(result.data)) return result.data;
  if (Array.isArray(result)) return result;
  if (result.results && typeof result.results === 'object') return Object.values(result.results);
  return [];
}

function getDocumentBatchRecordKeys(record, fallbackUrn = '') {
  return [
    record?.urn,
    record?.itemUrn,
    record?.lineageUrn,
    record?.versionUrn,
    record?.versionedFileUrn,
    record?.id,
    record?.itemId,
    record?.versionId,
    record?.attributes?.urn,
    record?.attributes?.itemUrn,
    record?.attributes?.lineageUrn,
    record?.attributes?.versionUrn,
    fallbackUrn
  ]
    .map((value) => String(value || '').trim())
    .filter(Boolean);
}

function getDocumentBatchAttributeSources(record) {
  return [
    record,
    record?.attributes || {},
    record?.customAttributes || [],
    record?.custom_attributes || [],
    record?.attributes?.customAttributes || [],
    record?.attributes?.custom_attributes || [],
    record?.attributes?.extension?.data || {},
    record?.extension?.data || {}
  ];
}

function hasUsefulDocumentBatchAttributes(record) {
  const sources = getDocumentBatchAttributeSources(record);
  return Boolean(
    readDocumentAttributeFromSources(sources, ['grd'], { exact: true }) ||
      readDocumentAttributeFromSources(sources, ['rev.', 'rev', 'revisao', 'revisão'], {
        exact: true,
        rejectValue: (value) => isInternalDocumentReviewValue(value)
      })
  );
}

async function requestDocumentBatchAttributes(req, projectGuid, urns) {
  const recordsByUrn = new Map();
  const chunkSize = 50;

  for (let index = 0; index < urns.length; index += chunkSize) {
    const chunk = urns.slice(index, index + chunkSize);
    const result = await callApsApi(req, `/bim360/docs/v1/projects/${encodeURIComponent(projectGuid)}/versions:batch-get`, {
      method: 'POST',
      body: { urns: chunk }
    });
    const records = collectDocumentBatchRecords(result);

    records.forEach((record, recordIndex) => {
      const fallbackUrn = chunk[recordIndex] || '';
      for (const key of getDocumentBatchRecordKeys(record, fallbackUrn)) {
        recordsByUrn.set(key, record);
      }
    });
  }

  return recordsByUrn;
}

async function getDocumentBatchAttributes(req, projectId, documents) {
  const projectGuid = getAccProjectId(projectId);
  const itemUrns = [...new Set(documents.map((document) => document.id).filter(Boolean))];
  let recordsByUrn = await requestDocumentBatchAttributes(req, projectGuid, itemUrns);

  if ([...recordsByUrn.values()].some(hasUsefulDocumentBatchAttributes)) {
    return recordsByUrn;
  }

  const versionPairs = await Promise.all(
    documents.slice(0, 120).map(async (document) => ({
      document,
      version: await getItemLatestVersion(req, projectId, document.id).catch(() => null)
    }))
  );
  const versionUrns = [...new Set(versionPairs.map(({ version }) => version?.id).filter(Boolean))];

  if (!versionUrns.length) {
    return recordsByUrn;
  }

  const versionRecordsByUrn = await requestDocumentBatchAttributes(req, projectGuid, versionUrns);
  for (const { document, version } of versionPairs) {
    const record = versionRecordsByUrn.get(version?.id);
    if (record) {
      recordsByUrn.set(document.id, record);
      recordsByUrn.set(version.id, record);
    }
  }

  return recordsByUrn;
}

async function enrichDocumentsWithBatchAttributes(req, projectId, documents) {
  if (!documents.length) return documents;

  let recordsByUrn;
  try {
    recordsByUrn = await getDocumentBatchAttributes(req, projectId, documents);
  } catch (error) {
    console.warn('Nao foi possivel ler atributos customizados dos documentos em lote.', error.message);
    return documents;
  }

  return documents.map((document) => {
    const batchRecord = recordsByUrn.get(document.id);
    if (!batchRecord) return document;

    const batchDocument = mapDocumentItem(
      { id: document.id, attributes: { displayName: document.name } },
      {
        id: batchRecord.versionUrn || batchRecord.versionedFileUrn || batchRecord.urn || document.id,
        attributes: {
          name: batchRecord.name || batchRecord.title || document.name,
          versionNumber: batchRecord.versionNumber || batchRecord.version || batchRecord.attributes?.versionNumber,
          lastModifiedTime: batchRecord.updateTime || batchRecord.updatedAt || batchRecord.attributes?.lastModifiedTime || document.updatedAt,
          createTime: batchRecord.createTime || batchRecord.createdAt || batchRecord.attributes?.createTime || document.versionCreatedAt
        },
        relationships: batchRecord.relationships || {},
        links: batchRecord.links || {}
      },
      document.folderPath,
      [
        ...getDocumentBatchAttributeSources(batchRecord)
      ]
    );

    return {
      ...document,
      ...batchDocument,
      name: document.name,
      folderPath: document.folderPath,
      description: batchDocument.description || document.description,
      grd: batchDocument.grd || document.grd,
      revision: batchDocument.revision || document.revision,
      version: batchDocument.version || document.version,
      emissionDate: batchDocument.emissionDate || document.emissionDate,
      updatedAt: batchDocument.updatedAt || document.updatedAt,
      versionCreatedAt: batchDocument.versionCreatedAt || document.versionCreatedAt,
      webView: batchDocument.webView || document.webView
    };
  });
}

function mapDocumentItem(item, version, folderPath, inheritedAttributeSources = []) {
  const itemAttributes = item.attributes || {};
  const versionAttributes = version?.attributes || {};
  const fileName = versionAttributes.name || itemAttributes.displayName || itemAttributes.name || item.id;
  const attributeSources = [
    version,
    item,
    versionAttributes,
    itemAttributes,
    versionAttributes.extension,
    versionAttributes.extension?.data,
    itemAttributes.extension,
    itemAttributes.extension?.data,
    version?.relationships,
    item.relationships,
    version?.meta,
    item.meta,
    ...inheritedAttributeSources
  ];
  const customVersionCandidate = readDocumentAttributeFromSources(attributeSources, [
    'versao',
    'versão',
    'version',
    'versao do documento',
    'versão do documento',
    'versao acc',
    'versão acc'
  ]);
  const customDescription =
    readDocumentAttributeFromSources(attributeSources, [
      'description',
      'descricao',
      'descrição',
      'descricao do documento',
      'descrição do documento'
    ]) ||
    readDocumentAttribute(version, ['description', 'descricao', 'descrição']) ||
    '';
  const customGrd =
    readDocumentAttributeFromSources(attributeSources, [
      'grd',
      'grd emissao',
      'grd emissão',
      'numero grd',
      'número grd',
      'guia',
      'remessa',
      'guia de remessa',
      'guia de remessa de documentos',
      'transmittal'
    ], { exact: true }) ||
    inferDocumentGrdFromText(
      readDocumentAttributeFromSources(attributeSources, ['grd'], { exact: false }) || ''
    ) ||
    '';
  const customRevision =
    readDocumentAttributeFromSources(attributeSources, [
      'rev.',
      'rev',
      'revisao',
      'revisão',
      'revisao do documento',
      'revisão do documento'
    ], {
      exact: true,
      rejectValue: (value) => isInternalDocumentReviewValue(value),
      excludeKeys: ['status da revisao', 'status da revisão', 'review status', 'review state']
    }) || '';
  const customEmissionDate = readDocumentAttributeFromSources(attributeSources, [
    'data emissao',
    'data emissão',
    'data de emissao',
    'data de emissão',
    'emissao',
    'emissão'
  ]);
  const versionNumber =
    readDocumentAttribute(version, ['versao', 'versão', 'version']) || versionAttributes.versionNumber || versionAttributes.version || null;
  const inferredRevision = inferDocumentRevisionFromName(fileName);
  const inferredGrd = inferDocumentGrdFromText(`${fileName} ${folderPath}`);

  return {
    id: item.id,
    name: fileName,
    folderPath,
    description: customDescription,
    grd: customGrd || inferredGrd,
    revision: customRevision || inferredRevision,
    version: (customVersionCandidate || versionNumber) ? String(customVersionCandidate || versionNumber).replace(/^v/i, 'V') : '',
    emissionDate: customEmissionDate || '',
    updatedAt: versionAttributes.lastModifiedTime || itemAttributes.lastModifiedTime || itemAttributes.createTime || null,
    versionCreatedAt: versionAttributes.createTime || versionAttributes.lastModifiedTime || itemAttributes.createTime || null,
    updatedBy:
      versionAttributes.lastModifiedUserName ||
      versionAttributes.lastModifiedUserId ||
      itemAttributes.lastModifiedUserName ||
      itemAttributes.createUserName ||
      '',
    fileType: fileName.includes('.') ? fileName.split('.').pop().toUpperCase() : '',
    webView: version?.links?.webView?.href || item.links?.webView?.href || null
  };
}


function normalizeDocumentText(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function normalizeProjectId(projectId) {
  return getAccProjectId(String(projectId || ''));
}

function cleanFolderLabel(value) {
  return String(value || '').replace(/\s+/g, ' ').trim();
}

function matchesDocumentListFolderName(value) {
  const normalizedName = normalizeFolderName(value);

  // Evita capturar pastas genéricas como "Teste LD Padrão" antes da pasta oficial.
  // A pasta correta do módulo Documentos Emitidos deve ser a pasta da Lista de Documentos
  // dentro de 05. DIRETRIZES DE PROJETO, não qualquer pasta que contenha a sigla LD.
  const isExactLdFolder = /^(\d{1,2}[. -]*)?ld$/i.test(normalizedName);

  return (
    normalizedName.includes('lista de documentos') ||
    normalizedName.includes('lista documentos') ||
    normalizedName.includes('document list') ||
    isExactLdFolder
  );
}


function isFolderEntry(entry) {
  return entry?.type === 'folders';
}

function getEntryName(entry) {
  return entry?.attributes?.displayName || entry?.attributes?.name || entry?.name || '';
}

function toFolderInfo(entry, pathParts = []) {
  const name = getEntryName(entry);
  const cleanPath = [...pathParts, name].filter(Boolean).map(cleanFolderLabel).join(' / ');
  return {
    id: entry.id,
    name,
    path: cleanPath || cleanFolderLabel(name),
    attributes: entry.attributes || {}
  };
}

function matchesProjectFilesFolderName(value) {
  const normalizedName = normalizeFolderName(value);
  return (
    normalizedName === 'project files' ||
    normalizedName === 'arquivos de projeto' ||
    normalizedName.includes('project files') ||
    normalizedName.includes('arquivos de projeto')
  );
}

function matchesDiretrizesFolderName(value) {
  const normalizedName = normalizeFolderName(value);
  return (
    normalizedName.includes('diretrizes de projeto') ||
    normalizedName.includes('05 diretrizes') ||
    normalizedName.includes('diretrizes')
  );
}

async function getFolderEntriesSafe(req, projectId, folderId) {
  const contents = await getFolderContents(req, projectId, folderId);
  return contents.data || [];
}

async function findPreferredPublishedFolder(req, projectId, topFolders = []) {
  // Caminho oficial do módulo Documentos Emitidos:
  // Project Files / 04. PUBLICADOS (DocEmit)
  // Esta busca direta evita varrer todas as pastas do projeto e reduz chamadas à Autodesk.
  const projectFilesFolder = topFolders.find((folder) => matchesProjectFilesFolderName(getEntryName(folder))) || topFolders[0];
  if (!projectFilesFolder) return null;

  const rootName = getEntryName(projectFilesFolder);
  const rootEntries = await getFolderEntriesSafe(req, projectId, projectFilesFolder.id);
  const publishedFolder = rootEntries.find((entry) => isFolderEntry(entry) && matchesPublishedFolderName(getEntryName(entry)));

  return publishedFolder ? toFolderInfo(publishedFolder, [rootName]) : null;
}

async function findPreferredDocumentListFolder(req, projectId, topFolders = []) {
  // Caminho oficial do módulo Documentos Emitidos:
  // Project Files / 05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS
  // Esta busca direta evita capturar pastas como "Teste LD Padrão" e reduz chamadas à Autodesk.
  const projectFilesFolder = topFolders.find((folder) => matchesProjectFilesFolderName(getEntryName(folder))) || topFolders[0];
  if (!projectFilesFolder) return null;

  const rootName = getEntryName(projectFilesFolder);
  const rootEntries = await getFolderEntriesSafe(req, projectId, projectFilesFolder.id);
  const diretrizesFolder = rootEntries.find((entry) => isFolderEntry(entry) && matchesDiretrizesFolderName(getEntryName(entry)));

  if (diretrizesFolder) {
    const diretrizesName = getEntryName(diretrizesFolder);
    const diretrizesEntries = await getFolderEntriesSafe(req, projectId, diretrizesFolder.id);
    const documentListFolder = diretrizesEntries.find((entry) => isFolderEntry(entry) && matchesDocumentListFolderName(getEntryName(entry)));

    if (documentListFolder) {
      return toFolderInfo(documentListFolder, [rootName, diretrizesName]);
    }
  }

  // Fallback controlado: caso a lista esteja temporariamente direto na raiz do Project Files.
  const directDocumentListFolder = rootEntries.find((entry) => isFolderEntry(entry) && matchesDocumentListFolderName(getEntryName(entry)));
  return directDocumentListFolder ? toFolderInfo(directDocumentListFolder, [rootName]) : null;
}

function isSpreadsheetEntry(entry) {
  const name = entry?.attributes?.displayName || entry?.attributes?.name || entry?.name || '';
  const extension = String(entry?.attributes?.extension?.data?.fileType || entry?.attributes?.extension || name.split('.').pop() || '').toLowerCase();
  return /\.(xlsx|xlsm|xls|csv)$/i.test(name) || ['xlsx', 'xlsm', 'xls', 'csv'].includes(extension);
}

function isPdfEntry(entry, pattern = '') {
  const attributes = entry?.attributes || {};
  const name = attributes.displayName || attributes.name || entry?.name || '';
  const extension = String(attributes.extension?.data?.fileType || attributes.extension || name.split('.').pop() || '').toLowerCase();
  const mimeType = String(attributes.mimeType || attributes.mime_type || '').toLowerCase();
  const isPdf = extension === 'pdf' || mimeType.includes('pdf') || /\.pdf$/i.test(name);
  if (!isPdf) return false;

  const normalizedPattern = normalizeDocumentText(pattern);
  return !normalizedPattern || normalizeDocumentText(name).includes(normalizedPattern);
}

async function findFolderByMatcher(req, projectId, folderId, matcher, pathParts = [], depth = 0) {
  if (depth > 8) return null;

  const contents = await getFolderContents(req, projectId, folderId);
  for (const entry of contents.data || []) {
    if (entry.type !== 'folders') continue;
    const folderName = entry.attributes?.displayName || entry.attributes?.name || '';
    const nextPath = [...pathParts, folderName];

    if (matcher(folderName)) {
      return {
        id: entry.id,
        name: folderName,
        path: nextPath.map(cleanFolderLabel).join(' / '),
        attributes: entry.attributes || {}
      };
    }

    const nestedFolder = await findFolderByMatcher(req, projectId, entry.id, matcher, nextPath, depth + 1).catch(() => null);
    if (nestedFolder) return nestedFolder;
  }

  return null;
}

async function findPublishedFolder(req, projectId, folderId, pathParts = [], depth = 0) {
  return findFolderByMatcher(req, projectId, folderId, matchesPublishedFolderName, pathParts, depth);
}

async function findDocumentListFolder(req, projectId, topFolders = []) {
  for (const folder of topFolders) {
    const folderName = folder.attributes?.displayName || folder.attributes?.name || '';
    if (matchesDocumentListFolderName(folderName)) {
      return {
        id: folder.id,
        name: folderName,
        path: cleanFolderLabel(folderName),
        attributes: folder.attributes || {}
      };
    }
  }

  for (const folder of topFolders) {
    const folderName = folder.attributes?.displayName || folder.attributes?.name || '';
    const nestedFolder = await findFolderByMatcher(req, projectId, folder.id, matchesDocumentListFolderName, [folderName]).catch(() => null);
    if (nestedFolder) return nestedFolder;
  }

  return null;
}

async function findDocumentListSpreadsheet(req, projectId, folderId, pathParts = [], depth = 0) {
  if (depth > 6) return null;

  const contents = await getFolderContents(req, projectId, folderId);
  const entries = contents.data || [];
  const spreadsheetItems = entries
    .filter((entry) => entry.type === 'items' && isSpreadsheetEntry(entry))
    .map((entry) => ({
      entry,
      name: entry.attributes?.displayName || entry.attributes?.name || entry.id,
      score: normalizeDocumentText(entry.attributes?.displayName || entry.attributes?.name || '').includes('lista') ? 2 : 1
    }))
    .sort((first, second) => second.score - first.score || String(first.name).localeCompare(String(second.name), 'pt-BR', { sensitivity: 'base' }));

  if (spreadsheetItems.length) {
    const selected = spreadsheetItems[0].entry;
    const version = await getItemLatestVersion(req, projectId, selected.id).catch(() => null);
    const name = selected.attributes?.displayName || selected.attributes?.name || selected.id;
    return {
      id: selected.id,
      name,
      item: selected,
      version,
      path: pathParts.map(cleanFolderLabel).join(' / '),
      updatedAt: version?.attributes?.lastModifiedTime || selected.attributes?.lastModifiedTime || selected.attributes?.createTime || null
    };
  }

  for (const folder of entries.filter((entry) => entry.type === 'folders')) {
    const folderName = folder.attributes?.displayName || folder.attributes?.name || '';
    const nestedSpreadsheet = await findDocumentListSpreadsheet(req, projectId, folder.id, [...pathParts, folderName], depth + 1).catch(() => null);
    if (nestedSpreadsheet) return nestedSpreadsheet;
  }

  return null;
}

async function listDocumentsInFolder(req, projectId, folderId, pathParts = [], documents = [], depth = 0, options = {}) {
  const maxDepth = options.maxDepth ?? 6;
  const maxDocuments = options.maxDocuments ?? 200;
  const includeVersions = options.includeVersions ?? true;
  const startedAt = options.startedAt ?? Date.now();
  const maxMs = options.maxMs ?? 15000;

  if (depth > maxDepth || documents.length >= maxDocuments || Date.now() - startedAt > maxMs) {
    return documents;
  }

  const contents = await getFolderContents(req, projectId, folderId);
  for (const entry of contents.data || []) {
    if (documents.length >= maxDocuments || Date.now() - startedAt > maxMs) break;

    const entryName = entry.attributes?.displayName || entry.attributes?.name || entry.id;
    if (entry.type === 'folders') {
      await listDocumentsInFolder(req, projectId, entry.id, [...pathParts, entryName], documents, depth + 1, options);
      continue;
    }

    if (entry.type !== 'items') continue;

    const version = includeVersions ? await getItemLatestVersion(req, projectId, entry.id).catch(() => null) : null;
    documents.push(mapDocumentItem(entry, version, pathParts.map(cleanFolderLabel).join(' / ')));
  }

  return documents;
}

function parseOssStorageId(storageId) {
  const value = String(storageId || '');
  const match = value.match(/urn:adsk\.objects:os\.object:([^/]+)\/(.+)$/i);
  if (!match) return { bucket: '', objectKey: '' };
  return {
    bucket: match[1],
    objectKey: match[2]
  };
}

async function downloadVersionBuffer(req, version) {
  const storageId = version?.relationships?.storage?.data?.id;
  const { bucket, objectKey } = parseOssStorageId(storageId);
  if (!bucket || !objectKey) {
    throw new Error('Versão sem storageId válido para download.');
  }

  const signedDownload = await callApsApi(
    req,
    `/oss/v2/buckets/${encodeURIComponent(bucket)}/objects/${encodeURIComponent(objectKey)}/signeds3download?minutesExpiration=10`
  );
  const downloadUrl = signedDownload.url || signedDownload.signedUrl || signedDownload.signed_url;
  if (!downloadUrl) {
    throw new Error('A Autodesk não retornou URL assinada para download.');
  }

  const response = await fetchWithRetry(downloadUrl, { method: 'GET' });
  if (!response.ok) {
    throw new Error(`Não foi possível baixar o arquivo da Autodesk. HTTP ${response.status}`);
  }

  return Buffer.from(await response.arrayBuffer());
}

function pickDocumentColumn(row, aliases) {
  return normalizeCellValue(findColumnValue(row, aliases));
}

function inferDocumentCodeFromName(value) {
  const text = String(value || '');
  const match = text.match(/\b[A-Z0-9]{2,}(?:[-_][A-Z0-9]{1,}){2,}\b/i);
  return match ? match[0].replace(/_/g, '-').toUpperCase() : '';
}

function parseDocumentListWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const centralG5SheetName = workbook.SheetNames.find((name) => normalizeSpreadsheetKey(name) === 'central g5');
  const sheetName =
    centralG5SheetName ||
    workbook.SheetNames.find((name) => normalizeSpreadsheetKey(name).includes('ld') || normalizeSpreadsheetKey(name).includes('lista')) ||
    workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const tableRows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false });

  const knownAliases = [
    'codigo',
    'código',
    'codigo g5engenharia',
    'código g5engenharia',
    'documento',
    'descricao',
    'descrição',
    'disciplina',
    'tipo',
    'titulo',
    'título',
    'status',
    'data de emissao',
    'data de emissão'
  ];
  const headerCandidate = tableRows
    .map((row, index) => ({
      index,
      score: row.reduce((sum, cell) => {
        const normalized = normalizeSpreadsheetKey(cell);
        if (!normalized) return sum;
        return sum + (knownAliases.some((alias) => normalized.includes(normalizeSpreadsheetKey(alias))) ? 1 : 0);
      }, 0)
    }))
    .reduce((best, candidate) => (candidate.score > best.score ? candidate : best), { index: 0, score: 0 });

  const headerRowIndex = headerCandidate.index;
  const headers = (tableRows[headerRowIndex] || []).map((header, index) => {
    const label = String(header || '').trim();
    return label ? `${label}__${index}` : `Coluna ${index + 1}`;
  });

  const rows = tableRows
    .slice(headerRowIndex + 1)
    .map((row, rowIndex) => {
      const raw = headers.reduce((record, header, index) => {
        record[header] = row[index] ?? '';
        return record;
      }, {});
      const code = pickDocumentColumn(raw, [
        'codigo g5engenharia',
        'código g5engenharia',
        'codigo g5 engenharia',
        'código g5 engenharia',
        'codigo de engenharia',
        'código de engenharia',
        'codigo',
        'código',
        'documento',
        'cod documento'
      ]);
      const title = pickDocumentColumn(raw, ['descricao', 'descrição', 'titulo', 'título', 'nome do documento', 'documento']);
      const discipline = pickDocumentColumn(raw, ['disciplina', 'área', 'area']);
      const documentType = pickDocumentColumn(raw, ['tipo de documento', 'tipo documento', 'tipo', 'classe']);
      const status = pickDocumentColumn(raw, ['status', 'situacao', 'situação']);
      const emissionDate = pickDocumentColumn(raw, ['data de emissao', 'data de emissão', 'emissao', 'emissão']);

      return {
        id: `ld-${rowIndex + 1}`,
        rowNumber: headerRowIndex + rowIndex + 2,
        code: code || inferDocumentCodeFromName(title),
        title,
        discipline,
        documentType,
        status,
        emissionDate,
        sourceSheet: sheetName,
        raw
      };
    })
    .filter((row) => row.code || row.title || row.discipline || row.documentType);

  return {
    sheetName,
    headerRowIndex,
    columns: headers.map((header) => header.replace(/__\d+$/, '')),
    rows
  };
}

async function enrichPublishedDocuments(req, projectId, documents) {
  return documents;
}

function compareDocumentListWithPublished(workbookData, documents) {
  const normalizedDocuments = documents.map((document) => ({
    document,
    key: normalizeSpreadsheetKey(`${document.name} ${document.description} ${document.folderPath}`),
    code: normalizeSpreadsheetKey(document.code || inferDocumentCodeFromName(document.name))
  }));

  return (workbookData.rows || []).map((row) => {
    const rowCode = normalizeSpreadsheetKey(row.code || inferDocumentCodeFromName(row.title));
    const rowTitle = normalizeSpreadsheetKey(row.title);
    const matched = normalizedDocuments.find(({ key, code }) =>
      (rowCode && (code === rowCode || key.includes(rowCode))) ||
      (rowTitle && key.includes(rowTitle))
    )?.document || null;
    const emittedDate = matched?.emissionDate || matched?.updatedAt || matched?.versionCreatedAt || row.emissionDate || '';

    return {
      ...row,
      originalStatus: row.status || '',
      status: matched ? 'Emitido' : 'Pendente',
      emitted: Boolean(matched),
      emittedFileName: matched?.name || '',
      emittedFolder: matched?.folderPath || '',
      emittedRevision: matched?.revision || '',
      emittedVersion: matched?.version || '',
      emittedGrd: matched?.grd || '',
      emittedDate,
      emissionDate: emittedDate,
      revision: matched?.revision || '',
      grd: matched?.grd || '',
      version: matched?.version || '',
      updatedAt: matched?.updatedAt || '',
      versionCreatedAt: matched?.versionCreatedAt || '',
      webView: matched?.webView || '',
      matchedDocumentId: matched?.id || ''
    };
  });
}


function getLatestDocumentComparisonCache(hubId, projectId, preferredMaxDocuments = 100) {
  const exactKey = `${hubId}:${projectId}:${preferredMaxDocuments}`;
  const exactEntry = documentComparisonCache.get(exactKey);
  if (exactEntry?.payload?.rows?.length) return { key: exactKey, ...exactEntry };

  let latest = null;
  for (const [key, entry] of documentComparisonCache.entries()) {
    if (!key.startsWith(`${hubId}:${projectId}:`)) continue;
    if (!entry?.payload?.rows?.length) continue;
    if (!latest || entry.cachedAt > latest.cachedAt) latest = { key, ...entry };
  }

  return latest;
}

function getExcelCellText(value) {
  if (value === null || value === undefined) return '';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value !== 'object') return String(value);
  if (Array.isArray(value.richText)) return value.richText.map((part) => part.text || '').join('');
  if (value.text) return String(value.text);
  if (value.result !== undefined) return String(value.result ?? '');
  if (value.hyperlink && value.text) return String(value.text);
  return String(value);
}

function findWorkbookWorksheet(workbook, preferredSheetName = '') {
  const normalizedPreferred = normalizeSpreadsheetKey(preferredSheetName);
  if (preferredSheetName) {
    const exact = workbook.getWorksheet(preferredSheetName);
    if (exact) return exact;
  }

  if (normalizedPreferred) {
    const byNormalizedName = workbook.worksheets.find((worksheet) => normalizeSpreadsheetKey(worksheet.name) === normalizedPreferred);
    if (byNormalizedName) return byNormalizedName;
  }

  return (
    workbook.worksheets.find((worksheet) => normalizeSpreadsheetKey(worksheet.name) === 'central g5') ||
    workbook.worksheets.find((worksheet) => normalizeSpreadsheetKey(worksheet.name).includes('ld') || normalizeSpreadsheetKey(worksheet.name).includes('lista')) ||
    workbook.worksheets[0]
  );
}

function findExcelDocumentColumns(worksheet) {
  const maxHeaderRows = Math.min(Math.max(worksheet.rowCount, 1), 30);
  let best = { rowNumber: 1, score: -1, statusColumn: 0, emissionDateColumn: 0, lastColumn: 1 };

  for (let rowNumber = 1; rowNumber <= maxHeaderRows; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    let score = 0;
    let statusColumn = 0;
    let emissionDateColumn = 0;
    let lastColumn = 1;

    row.eachCell({ includeEmpty: false }, (cell, columnNumber) => {
      const normalized = normalizeSpreadsheetKey(getExcelCellText(cell.value || cell.text));
      if (!normalized) return;
      lastColumn = Math.max(lastColumn, columnNumber);

      if (['status', 'situacao', 'situação'].includes(normalized)) {
        statusColumn = columnNumber;
        score += 4;
      }

      if (['data de emissao', 'data de emissão', 'data emissao', 'data emissão', 'emissao', 'emissão'].includes(normalized)) {
        emissionDateColumn = columnNumber;
        score += 4;
      }

      if (
        normalized.includes('codigo') ||
        normalized.includes('código') ||
        normalized.includes('disciplina') ||
        normalized.includes('tipo de documento') ||
        normalized.includes('descricao') ||
        normalized.includes('descrição')
      ) {
        score += 1;
      }
    });

    if (score > best.score) best = { rowNumber, score, statusColumn, emissionDateColumn, lastColumn };
  }

  return best;
}

function ensureExcelColumn(worksheet, headerRowNumber, preferredColumn, title) {
  if (preferredColumn) return preferredColumn;
  const headerRow = worksheet.getRow(headerRowNumber);
  const nextColumn = Math.max(headerRow.actualCellCount + 1, headerRow.cellCount + 1, 1);
  const headerCell = headerRow.getCell(nextColumn);
  headerCell.value = title;

  const previousHeaderCell = nextColumn > 1 ? headerRow.getCell(nextColumn - 1) : null;
  if (previousHeaderCell?.style) headerCell.style = { ...previousHeaderCell.style };

  return nextColumn;
}

function parseExcelDateValue(value) {
  const isoDate = parseSpreadsheetDate(value);
  if (!isoDate) return null;
  const [year, month, day] = isoDate.split('-').map(Number);
  if (!year || !month || !day) return null;
  return new Date(year, month - 1, day);
}

async function updateDocumentWorkbookBuffer(originalBuffer, comparisonRows = [], preferredSheetName = '') {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(originalBuffer);

  const worksheet = findWorkbookWorksheet(workbook, preferredSheetName);
  if (!worksheet) {
    throw new Error('Nao encontrei a aba Central G5 ou outra aba valida na planilha.');
  }

  const columns = findExcelDocumentColumns(worksheet);
  const statusColumn = ensureExcelColumn(worksheet, columns.rowNumber, columns.statusColumn, 'Status');
  const emissionDateColumn = ensureExcelColumn(
    worksheet,
    columns.rowNumber,
    columns.emissionDateColumn || (statusColumn === columns.lastColumn + 1 ? 0 : columns.emissionDateColumn),
    'Data de Emissão'
  );
  const headerStatusCell = worksheet.getRow(columns.rowNumber).getCell(statusColumn);
  const headerDateCell = worksheet.getRow(columns.rowNumber).getCell(emissionDateColumn);
  if (!headerStatusCell.value) headerStatusCell.value = 'Status';
  if (!headerDateCell.value) headerDateCell.value = 'Data de Emissão';

  let updatedRows = 0;
  let emittedRows = 0;
  let pendingRows = 0;

  for (const comparisonRow of comparisonRows) {
    const rowNumber = Number(comparisonRow.rowNumber);
    if (!Number.isFinite(rowNumber) || rowNumber <= columns.rowNumber) continue;

    const row = worksheet.getRow(rowNumber);
    const status = comparisonRow.emitted ? 'Emitido' : 'Pendente';
    const statusCell = row.getCell(statusColumn);
    const dateCell = row.getCell(emissionDateColumn);

    statusCell.value = status;
    if (comparisonRow.emitted) {
      emittedRows += 1;
      const emittedDate = parseExcelDateValue(
        comparisonRow.emittedDate || comparisonRow.updatedAt || comparisonRow.versionCreatedAt || comparisonRow.emissionDate
      );
      dateCell.value = emittedDate || '';
      if (emittedDate && !dateCell.numFmt) dateCell.numFmt = 'dd/mm/yyyy';
    } else {
      pendingRows += 1;
      dateCell.value = '';
    }

    row.commit?.();
    updatedRows += 1;
  }

  const outputBuffer = await workbook.xlsx.writeBuffer();
  return {
    buffer: Buffer.from(outputBuffer),
    updatedRows,
    emittedRows,
    pendingRows,
    sheetName: worksheet.name
  };
}

async function createStorageLocation(req, projectId, folderId, fileName) {
  try {
    const result = await callApsApi(req, `/data/v1/projects/${encodeURIComponent(projectId)}/storage`, {
      method: 'POST',
      body: {
        jsonapi: { version: '1.0' },
        data: {
          type: 'objects',
          attributes: { name: fileName },
          relationships: {
            target: {
              data: {
                type: 'folders',
                id: folderId
              }
            }
          }
        }
      }
    });

    const storageId = result?.data?.id;
    if (!storageId) throw new Error('A Autodesk nao retornou o storageId para enviar a planilha atualizada.');
    return storageId;
  } catch (error) {
    error.apsStep = 'Criar area temporaria de upload no ACC';
    throw error;
  }
}

async function uploadBufferToStorage(req, storageId, buffer, mimeType) {
  const { bucket, objectKey } = parseOssStorageId(storageId);
  if (!bucket || !objectKey) throw new Error('StorageId invalido para upload da nova versao.');

  let signedUpload;
  try {
    signedUpload = await callApsApi(
      req,
      `/oss/v2/buckets/${encodeURIComponent(bucket)}/objects/${encodeURIComponent(objectKey)}/signeds3upload?minutesExpiration=10`
    );
  } catch (error) {
    error.apsStep = 'Gerar URL assinada para upload da planilha';
    throw error;
  }

  const uploadUrl =
    signedUpload.url ||
    signedUpload.signedUrl ||
    signedUpload.signed_url ||
    signedUpload.urls?.[0] ||
    signedUpload.uploadUrls?.[0] ||
    signedUpload.upload_url;
  const uploadKey = signedUpload.uploadKey || signedUpload.upload_key;

  if (!uploadUrl || !uploadKey) {
    throw new Error('A Autodesk nao retornou URL assinada completa para upload da planilha.');
  }

  const uploadResponse = await fetchWithRetry(uploadUrl, {
    method: 'PUT',
    headers: {
      'Content-Type': mimeType,
      'Content-Length': String(buffer.length)
    },
    body: buffer
  });

  if (!uploadResponse.ok) {
    const details = await uploadResponse.text().catch(() => '');
    const error = new Error(`Nao foi possivel enviar a planilha atualizada para a URL assinada. HTTP ${uploadResponse.status}. ${details}`);
    error.status = uploadResponse.status;
    error.apsStep = 'Enviar arquivo Excel para o storage temporario';
    throw error;
  }

  try {
    await callApsApi(req, `/oss/v2/buckets/${encodeURIComponent(bucket)}/objects/${encodeURIComponent(objectKey)}/signeds3upload`, {
      method: 'POST',
      body: { uploadKey }
    });
  } catch (error) {
    error.apsStep = 'Confirmar upload S3 na Autodesk';
    throw error;
  }
}

async function createNewItemVersion(req, projectId, itemId, storageId, fileName) {
  try {
    return await callApsApi(req, `/data/v1/projects/${encodeURIComponent(projectId)}/versions`, {
      method: 'POST',
      body: {
        jsonapi: { version: '1.0' },
        data: {
          type: 'versions',
          attributes: {
            name: fileName,
            extension: {
              type: 'versions:autodesk.bim360:File',
              version: '1.0'
            }
          },
          relationships: {
            item: {
              data: {
                type: 'items',
                id: itemId
              }
            },
            storage: {
              data: {
                type: 'objects',
                id: storageId
              }
            }
          }
        }
      }
    });
  } catch (error) {
    error.apsStep = 'Criar nova versao do arquivo Excel no ACC';
    throw error;
  }
}

async function locateDocumentListSpreadsheet(req, hubId, projectId) {
  const topFoldersResult = await callApsApi(
    req,
    `/project/v1/hubs/${encodeURIComponent(hubId)}/projects/${encodeURIComponent(projectId)}/topFolders`
  );
  const topFolders = topFoldersResult.data || [];
  const preferredDocumentListFolder = await findPreferredDocumentListFolder(req, projectId, topFolders).catch(() => null);
  let documentListFolder = preferredDocumentListFolder || topFolders.find((folder) =>
    matchesDocumentListFolderName(folder.attributes?.displayName || folder.attributes?.name)
  );
  let documentListFolderPath = preferredDocumentListFolder?.path || cleanFolderLabel(
    documentListFolder?.attributes?.displayName || documentListFolder?.attributes?.name || '05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS'
  );

  if (!documentListFolder) {
    const nestedFolder = await findDocumentListFolder(req, projectId, topFolders).catch(() => null);
    if (nestedFolder) {
      documentListFolder = nestedFolder;
      documentListFolderPath = nestedFolder.path;
    }
  }

  if (!documentListFolder?.id) {
    throw new Error('Nao encontrei a pasta 05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS neste projeto.');
  }

  const spreadsheetItem = await findDocumentListSpreadsheet(req, projectId, documentListFolder.id, [documentListFolderPath]);
  if (!spreadsheetItem?.version) {
    throw new Error(`Encontrei a pasta "${documentListFolderPath}", mas nao encontrei uma planilha Excel valida nela.`);
  }

  return { documentListFolder, documentListFolderPath, spreadsheetItem };
}

function mapAttributeMappings(result) {
  const mappingsByDefinitionId = new Map();

  for (const mapping of (result.results || result.data || []).filter(isActiveConfiguration)) {
    const attributeDefinitionId =
      mapping.attributeDefinitionId ||
      mapping.issueAttributeDefinitionId ||
      mapping.customAttributeDefinitionId ||
      mapping.definitionId;
    const mappedItemId = mapping.mappedItemId || mapping.issueTypeId || mapping.issueSubtypeId || mapping.itemId;
    const mappedItemType =
      mapping.mappedItemType ||
      (mapping.issueSubtypeId ? 'issueSubtype' : mapping.issueTypeId ? 'issueType' : mapping.itemType);

    if (!attributeDefinitionId || !mappedItemId) continue;

    const mappings = mappingsByDefinitionId.get(attributeDefinitionId) || [];
    mappings.push({
      mappedItemId,
      mappedItemType
    });
    mappingsByDefinitionId.set(attributeDefinitionId, mappings);
  }

  return mappingsByDefinitionId;
}

async function getFolderContents(req, projectId, folderId) {
  const firstEndpoint = `/data/v1/projects/${encodeURIComponent(projectId)}/folders/${encodeURIComponent(folderId)}/contents?page[limit]=200`;
  const firstPage = await callApsApi(req, firstEndpoint);
  const allEntries = [...(firstPage.data || [])];
  let nextUrl = firstPage.links?.next?.href;
  let pageCount = 1;

  while (nextUrl && pageCount < 8) {
    const nextPage = await callApsUrl(req, nextUrl);
    allEntries.push(...(nextPage.data || []));
    nextUrl = nextPage.links?.next?.href;
    pageCount += 1;
  }

  return {
    ...firstPage,
    data: allEntries
  };
}

async function getItemLatestVersion(req, projectId, itemId) {
  const result = await callApsApi(
    req,
    `/data/v1/projects/${encodeURIComponent(projectId)}/items/${encodeURIComponent(itemId)}/versions?page[limit]=1`
  );

  return result.data?.[0] || null;
}

async function findFolderByName(req, projectId, folderId, targetName, pathParts = [], depth = 0) {
  if (depth > 8) {
    return null;
  }

  const contents = await getFolderContents(req, projectId, folderId);
  const normalizedTarget = targetName.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();

  for (const entry of contents.data || []) {
    if (entry.type !== 'folders') continue;

    const folderName = entry.attributes?.displayName || entry.attributes?.name || '';
    const normalizedFolderName = folderName.normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
    const nextPath = [...pathParts, folderName];

    if (normalizedFolderName.includes(normalizedTarget)) {
      return {
        id: entry.id,
        name: folderName,
        path: nextPath.join(' / ')
      };
    }

    const nestedFolder = await findFolderByName(req, projectId, entry.id, targetName, nextPath, depth + 1).catch(() => null);
    if (nestedFolder) {
      return nestedFolder;
    }
  }

  return null;
}

function normalizeFolderName(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}

function matchesPublishedFolderName(value) {
  const normalizedName = normalizeFolderName(value);

  return (
    normalizedName.includes('publicado') ||
    normalizedName.includes('publicados') ||
    normalizedName.includes('docemit') ||
    normalizedName.includes('emitidos') ||
    normalizedName.includes('04 publicados')
  );
}

function getDefinitionOptionLabelSafe(definition, value) {
  const options =
    definition?.metadata?.list?.options ||
    definition?.metadata?.options ||
    definition?.options ||
    definition?.permittedValues ||
    [];

  const option = options.find((item) => {
    const optionId = item.id || item.valueId || item.key;
    return optionId === value || item.value === value || item.label === value || item.name === value;
  });

  return option?.value || option?.label || option?.name || null;
}


app.get('/api/auth/login', async (req, res, next) => {
  try {
    const { clientId } = getApsCredentials();
    const state = crypto.randomBytes(16).toString('hex');
    req.session.oauthState = state;
    await saveSession(req);

    const params = new URLSearchParams({
      response_type: 'code',
      client_id: clientId,
      redirect_uri: callbackUrl,
      scope: apsScopes.join(' '),
      state
    });

    res.redirect(`${apsAuthorizeUrl}?${params.toString()}`);
  } catch (error) {
    next(error);
  }
});

app.get('/api/auth/callback', async (req, res, next) => {
  try {
    const { code, state, error } = req.query;

    if (error) {
      res.redirect(`${frontendUrl}?login=error`);
      return;
    }

    if (!code || typeof code !== 'string') {
      res.redirect(`${frontendUrl}?login=error`);
      return;
    }

    if (req.session.oauthState && state !== req.session.oauthState) {
      res.redirect(`${frontendUrl}?login=error`);
      return;
    }

    const token = await exchangeCodeForToken(code);
    req.session.apsToken = normalizeToken(token);
    delete req.session.oauthState;
    await saveSession(req);

    res.redirect(frontendUrl);
  } catch (error) {
    next(error);
  }
});

app.post('/api/auth/keep-alive', async (req, res, next) => {
  try {
    const accessToken = await getValidAccessToken(req);
    if (!accessToken) {
      res.status(401).json({ error: 'Sessão Autodesk não autenticada.' });
      return;
    }

    await saveSession(req);
    res.json({ authenticated: true });
  } catch (error) {
    next(error);
  }
});

app.post('/api/auth/logout', async (req, res) => {
  req.session.destroy(() => {
    res.clearCookie('g5_issues_planner_session');
    res.json({ ok: true });
  });
});

app.get('/api/me', async (req, res, next) => {
  try {
    const accessToken = await getValidAccessToken(req);
    const identity = readIdentityFromAccessToken(accessToken || '');

    res.json({
      name: identity.name || readNameFromAccessToken(accessToken) || 'Usuário Autodesk',
      email: identity.email || null,
      ids: identity.ids || []
    });
  } catch (error) {
    next(error);
  }
});

app.get('/api/auth/scopes', async (req, res, next) => {
  try {
    const accessToken = await getValidAccessToken(req);
    if (!accessToken) {
      res.status(401).json({ authenticated: false, scopes: [], required: apsScopes });
      return;
    }

    res.json({
      authenticated: true,
      requested: apsScopes,
      granted: getSessionTokenScopes(req),
      canWriteSpreadsheet: hasWriteTokenScopes(req)
    });
  } catch (error) {
    next(error);
  }
});

app.get('/api/hubs', async (req, res, next) => {
  try {
    // Hubs são as contas/ambientes ACC, Forma, BIM 360 ou Fusion acessíveis ao usuário.
    const result = await callApsApi(req, '/project/v1/hubs');
    const hubs = (result.data || []).map((hub) => ({
      id: hub.id,
      name: hub.attributes?.name || hub.id,
      type: hub.attributes?.extension?.type || hub.type
    }));

    res.json(hubs);
  } catch (error) {
    next(error);
  }
});

app.get('/api/hubs/:hubId/projects', async (req, res, next) => {
  try {
    // Projetos são carregados somente depois que um Hub é escolhido no frontend.
    const result = await callApsApi(req, `/project/v1/hubs/${encodeURIComponent(req.params.hubId)}/projects`);
    const candidateProjects = (result.data || [])
      .filter(isActiveProject)
      .filter((project) => !isTemplateProject(project));
    const businessUnitNamesById = await getBusinessUnitNamesById(req, getAccountIdFromHubId(req.params.hubId));
    const adminMetadataByProjectId = await getAdminProjectMetadataMap(req, req.params.hubId, candidateProjects);
    const projectsWithMembership = await Promise.all(
      candidateProjects.map(async (project) => ({
        project,
        isMember: await isCurrentUserProjectMember(req, project.id).catch(() => true)
      }))
    );
    const projects = projectsWithMembership
      .filter(({ isMember }) => isMember)
      .map(({ project }) => project)
      .map((project) => {
        const adminData = adminMetadataByProjectId.get(normalizeAccProjectId(project.id)) || {};
        const projectWithAdminData = { ...project, adminData };
        return {
          id: project.id,
          name: project.attributes?.name || project.id,
          type: project.attributes?.extension?.type || project.type,
          businessUnit: getProjectBusinessUnit(projectWithAdminData, businessUnitNamesById)
        };
      });

    res.json(projects);
  } catch (error) {
    next(error);
  }
});

app.get('/api/hubs/:hubId/projects/:projectId/published-documents', async (req, res, next) => {
  try {
    const { hubId, projectId } = req.params;
    const requestedLimit = Number(req.query.limit || 100);
    const maxDocuments = Math.min(Math.max(requestedLimit, 20), 120);
    const forceRefresh = ['1', 'true', 'yes'].includes(String(req.query.refresh || '').toLowerCase());
    const cacheKey = `${hubId}:${projectId}:${maxDocuments}`;
    const cachedComparison = documentComparisonCache.get(cacheKey);

    if (!forceRefresh && cachedComparison && Date.now() - cachedComparison.cachedAt < documentComparisonCacheTtlMs) {
      res.json({
        ...cachedComparison.payload,
        cached: true,
        cachedAt: cachedComparison.cachedAt
      });
      return;
    }

    const topFoldersResult = await callApsApi(
      req,
      `/project/v1/hubs/${encodeURIComponent(hubId)}/projects/${encodeURIComponent(projectId)}/topFolders`
    );

    const topFolders = topFoldersResult.data || [];

    const preferredPublishedFolder = await findPreferredPublishedFolder(req, projectId, topFolders).catch(() => null);
    const preferredDocumentListFolder = await findPreferredDocumentListFolder(req, projectId, topFolders).catch(() => null);

    let publishedFolder = preferredPublishedFolder || topFolders.find((folder) => matchesPublishedFolderName(folder.attributes?.displayName || folder.attributes?.name));
    let publishedFolderPath = preferredPublishedFolder?.path || cleanFolderLabel(publishedFolder?.attributes?.displayName || publishedFolder?.attributes?.name || '04. PUBLICADOS (DocEmit)');
    let documentListFolder = preferredDocumentListFolder || topFolders.find((folder) =>
      matchesDocumentListFolderName(folder.attributes?.displayName || folder.attributes?.name)
    );
    let documentListFolderPath = preferredDocumentListFolder?.path || cleanFolderLabel(
      documentListFolder?.attributes?.displayName || documentListFolder?.attributes?.name || '05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS'
    );

    if (!publishedFolder) {
      for (const folder of topFolders) {
        const folderName = folder.attributes?.displayName || folder.attributes?.name || '';
        const nestedFolder = await findPublishedFolder(req, projectId, folder.id, [folderName]).catch(() => null);
        if (nestedFolder) {
          publishedFolder = nestedFolder;
          publishedFolderPath = nestedFolder.path;
          break;
        }
      }
    }

    if (!documentListFolder) {
      const nestedFolder = await findDocumentListFolder(req, projectId, topFolders).catch(() => null);
      if (nestedFolder) {
        documentListFolder = nestedFolder;
        documentListFolderPath = nestedFolder.path;
      }
    }

    if (!publishedFolder) {
      res.json({
        folder: null,
        documents: [],
        spreadsheet: null,
        rows: [],
        message: 'Nao encontrei a pasta 04. PUBLICADOS (DocEmit) neste projeto.'
      });
      return;
    }

    const scanStartedAt = Date.now();
    const maxScanMs = 14000;
    let documents = await listDocumentsInFolder(req, projectId, publishedFolder.id, [publishedFolderPath], [], 0, {
      maxDepth: 7,
      maxDocuments,
      includeVersions: false,
      startedAt: scanStartedAt,
      maxMs: maxScanMs
    });
    // Evita chamadas extras de atributos/versões em lote, que podem estourar o limite diário da Autodesk.
    // Para o módulo Documentos Emitidos, a comparação precisa principalmente do nome/código do arquivo publicado.
    const scanWasLimited = documents.length >= maxDocuments || Date.now() - scanStartedAt > maxScanMs;
    let spreadsheet = null;
    let rows = [];
    let columns = [];
    let sheetName = '';
    let message = '';

    if (documentListFolder) {
      const spreadsheetItem = await findDocumentListSpreadsheet(req, projectId, documentListFolder.id, [documentListFolderPath]).catch(
        () => null
      );

      if (spreadsheetItem?.version) {
        try {
          const workbookBuffer = await downloadVersionBuffer(req, spreadsheetItem.version);
          const workbookData = parseDocumentListWorkbook(workbookBuffer);
          documents = await enrichPublishedDocuments(req, projectId, documents, workbookData.rows, maxDocuments);
          rows = compareDocumentListWithPublished(workbookData, documents);
          columns = workbookData.columns;
          sheetName = workbookData.sheetName;
          spreadsheet = {
            name: spreadsheetItem.name,
            path: spreadsheetItem.path || documentListFolderPath,
            itemId: spreadsheetItem.id,
            versionId: spreadsheetItem.version?.id || '',
            folderId: documentListFolder.id,
            updatedAt: spreadsheetItem.updatedAt,
            webView: spreadsheetItem.version?.links?.webView?.href || spreadsheetItem.item?.links?.webView?.href || null
          };
        } catch (spreadsheetError) {
          message = `Encontrei a planilha "${spreadsheetItem.name}" em "${spreadsheetItem.path || documentListFolderPath}", mas nao consegui baixar ou ler o arquivo. ${spreadsheetError.message}`;
        }
      } else {
        message = `Encontrei a pasta "${documentListFolderPath}", mas nao encontrei uma planilha Excel nela.`;
      }
    } else {
      message = 'Nao encontrei a pasta 05. DIRETRIZES DE PROJETO / 05. LISTA DE DOCUMENTOS neste projeto.';
    }

    const payload = {
      folder: {
        id: publishedFolder.id,
        name:
          publishedFolder.name ||
          cleanFolderLabel(publishedFolder.attributes?.displayName || publishedFolder.attributes?.name || '04. PUBLICADOS (DocEmit)'),
        path: publishedFolderPath
      },
      spreadsheet,
      sheetName,
      columns,
      rows,
      documents: documents.sort((firstDocument, secondDocument) =>
        firstDocument.name.localeCompare(secondDocument.name, 'pt-BR', { sensitivity: 'base' })
      ),
      partial: scanWasLimited,
      message:
        message ||
        (scanWasLimited
          ? 'A busca foi limitada para evitar queda do servidor. Se faltarem documentos, clique em Atualizar comparativo novamente.'
          : '')
    };

    documentComparisonCache.set(cacheKey, {
      cachedAt: Date.now(),
      payload
    });

    res.json({
      ...payload,
      cached: false,
      cachedAt: Date.now()
    });
  } catch (error) {
    console.error(error);

    const { hubId, projectId } = req.params;
    const requestedLimit = Number(req.query.limit || 100);
    const maxDocuments = Math.min(Math.max(requestedLimit, 20), 120);
    const cacheKey = `${hubId}:${projectId}:${maxDocuments}`;
    const cachedComparison = documentComparisonCache.get(cacheKey);

    if (cachedComparison && isApsQuotaLimitError(error)) {
      res.json({
        ...cachedComparison.payload,
        cached: true,
        cachedAt: cachedComparison.cachedAt,
        message:
          'A Autodesk atingiu o limite temporario de consultas. Mantive a ultima leitura salva deste projeto; aguarde alguns minutos antes de forcar nova atualizacao.'
      });
      return;
    }

    res.status(error.status || 500).json({
      folder: null,
      documents: [],
      spreadsheet: null,
      rows: [],
      message: isApsQuotaLimitError(error)
        ? 'A Autodesk atingiu o limite temporario de consultas deste app. Aguarde alguns minutos e atualize novamente. Para reduzir novas chamadas, evite clicar varias vezes em Atualizar dados.'
        : formatApsError(error.message || 'Nao foi possivel concluir a solicitacao.')
    });
  }
});


app.post('/api/hubs/:hubId/projects/:projectId/published-documents/update-spreadsheet', async (req, res, next) => {
  try {
    if (!hasWriteTokenScopes(req)) {
      res.status(401).json({
        message:
          'Sua sessao Autodesk atual nao possui data:create/data:write. Clique em Sair, entre novamente com a Autodesk e aceite as novas permissoes.' +
          formatTokenScopeHint(req)
      });
      return;
    }

    const { hubId, projectId } = req.params;
    const requestedLimit = Number(req.body?.limit || req.query?.limit || 100);
    const maxDocuments = Math.min(Math.max(requestedLimit, 20), 120);
    const cachedComparison = getLatestDocumentComparisonCache(hubId, projectId, maxDocuments);

    if (!cachedComparison?.payload?.rows?.length) {
      res.status(409).json({
        message:
          'Antes de atualizar a planilha no ACC, carregue o modulo Documentos Emitidos e clique em Atualizar dados uma vez para gerar o comparativo.'
      });
      return;
    }

    const comparisonRows = cachedComparison.payload.rows || [];
    const { documentListFolder, documentListFolderPath, spreadsheetItem } = await locateDocumentListSpreadsheet(req, hubId, projectId);
    const originalWorkbookBuffer = await downloadVersionBuffer(req, spreadsheetItem.version);
    const workbookUpdate = await updateDocumentWorkbookBuffer(
      originalWorkbookBuffer,
      comparisonRows,
      cachedComparison.payload.sheetName || comparisonRows[0]?.sourceSheet || 'Central G5'
    );

    if (!workbookUpdate.updatedRows) {
      res.status(409).json({ message: 'Nao encontrei linhas validas para atualizar na planilha.' });
      return;
    }

    const fileName = spreadsheetItem.name || 'Lista de Documentos.xlsx';
    const mimeType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const storageId = await createStorageLocation(req, projectId, documentListFolder.id, fileName);
    await uploadBufferToStorage(req, storageId, workbookUpdate.buffer, mimeType);
    const versionResult = await createNewItemVersion(req, projectId, spreadsheetItem.id, storageId, fileName);

    documentComparisonCache.delete(cachedComparison.key);

    res.json({
      success: true,
      message: `Planilha atualizada no ACC como nova versao. Foram atualizadas ${workbookUpdate.updatedRows} linhas na aba ${workbookUpdate.sheetName}.`,
      updatedRows: workbookUpdate.updatedRows,
      emittedRows: workbookUpdate.emittedRows,
      pendingRows: workbookUpdate.pendingRows,
      spreadsheet: {
        name: fileName,
        path: spreadsheetItem.path || documentListFolderPath,
        itemId: spreadsheetItem.id,
        versionId: versionResult?.data?.id || '',
        updatedAt: new Date().toISOString()
      }
    });
  } catch (error) {
    console.error(error);
    const message = String(error?.message || 'Nao foi possivel atualizar a planilha no ACC.');
    const missingScope =
      error?.status === 403 ||
      normalizeSpreadsheetKey(message).includes('forbidden') ||
      normalizeSpreadsheetKey(message).includes('insufficient scope') ||
      normalizeSpreadsheetKey(message).includes('scope');

    const stepText = error?.apsStep ? `Etapa: ${error.apsStep}. ` : '';
    const endpointText = error?.endpoint ? `Endpoint: ${error.endpoint}. ` : '';
    const scopeHint = formatTokenScopeHint(req);

    res.status(error.status || 500).json({
      message: missingScope
        ? `${stepText}${endpointText}A Autodesk recusou a gravacao da planilha. O app pediu data:create/data:write, mas a sessao atual ou esta etapa da API nao foi aceita. Saia do app, entre novamente com a Autodesk e tente de novo.${scopeHint}`
        : `${stepText}${endpointText}${formatApsError(message)}${scopeHint}`
    });
  }
});

app.get('/api/projects/:projectId/issues', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const requestedLimit = Number.parseInt(String(req.query.limit || '200'), 10);
    const safeLimit = Number.isFinite(requestedLimit) ? Math.min(Math.max(requestedLimit, 1), 200) : 200;
    const query = new URLSearchParams({
      limit: String(safeLimit)
    });

    // A Issues API do ACC trabalha por projeto.
    const [issuesResult, definitionsResult, issueTypesResult, projectUsersResult] = await Promise.all([
      fetchAllIssuePages(req, projectId, safeLimit),
      callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions?limit=200`).catch(() => ({
        results: []
      })),
      callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-types?include=subtypes&limit=200`).catch(() => ({
        results: []
      })),
      callApsApi(req, `/construction/admin/v1/projects/${projectId}/users?limit=200`).catch(() => ({
        results: []
      }))
    ]);
    const attributeDefinitions = mapAttributeDefinitions(definitionsResult);
    const issueTypeMaps = mapIssueTypes(issueTypesResult);
    const projectUsers = mapProjectUsers(projectUsersResult);
    const issues = (issuesResult.issues || []).map((issue) =>
      mapIssue(issue, attributeDefinitions, issueTypeMaps, projectUsers)
    );

    res.json({ issues, meta: { loaded: issues.length, pages: issuesResult.pages } });
  } catch (error) {
    next(error);
  }
});

app.get('/api/projects/:projectId/issues/:issueId/details', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const encodedIssueId = encodeURIComponent(req.params.issueId);

    const [issue, definitionsResult, issueTypesResult, projectUsersResult] = await Promise.all([
      callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues/${encodedIssueId}`),
      callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions?limit=200`).catch(() => ({
        results: []
      })),
      callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-types?include=subtypes&limit=200`).catch(() => ({
        results: []
      })),
      callApsApi(req, `/construction/admin/v1/projects/${projectId}/users?limit=200`).catch(() => ({
        results: []
      }))
    ]);

    const attributeDefinitions = mapAttributeDefinitions(definitionsResult);
    const issueTypeMaps = mapIssueTypes(issueTypesResult);
    const projectUsers = mapProjectUsers(projectUsersResult);
    const commentsResult = await getIssueComments(req, projectId, req.params.issueId, projectUsers);

    res.json({
      ...mapIssue(issue, attributeDefinitions, issueTypeMaps, projectUsers),
      comments: commentsResult.comments,
      commentsWarning: commentsResult.warning
    });
  } catch (error) {
    next(error);
  }
});

app.get('/api/projects/:projectId/cronograma/issue-report', async (req, res, next) => {
  try {
    const projectId = normalizeProjectId(req.params.projectId);
    const rootFolder = req.query.rootFolder || 'Project Files';
    const folderPathQuery = String(req.query.folderPath || '03. COORDENAÇÃO (BIM) > 01. Rotina de Qualidade');
    const filePattern = String(req.query.filePattern || 'Detalhe Issues');
    const topFoldersResult = await callApsApi(req, `/data/v1/projects/${encodeURIComponent(projectId)}/folders/${encodeURIComponent(rootFolder)}/contents`);
    const topFolders = (topFoldersResult.data || []).filter((entry) => entry.type === 'folders');
    const pathParts = folderPathQuery.split('>').map((part) => part.trim()).filter(Boolean);
    let currentFolder = null;
    for (const [index, part] of pathParts.entries()) {
      const sourceEntries = index === 0
        ? topFolders
        : (await getFolderContents(req, projectId, currentFolder.id)).data.filter((entry) => entry.type === 'folders');
      currentFolder = sourceEntries.map((entry) => ({
        id: entry.id,
        name: entry.attributes?.displayName || entry.attributes?.name || ''
      })).find((entry) => normalizeFolderName(entry.name).includes(normalizeFolderName(part)));
      if (!currentFolder) break;
    }

    if (!currentFolder?.id) {
      return res.status(404).json({ found: false, message: 'Pasta do relatório não encontrada.' });
    }

    const folderDocs = await listDocumentsInFolder(req, projectId, currentFolder.id, pathParts, [], 0, { maxDepth: 2, maxDocuments: 300, includeVersions: true });
    const pdfDocs = folderDocs.filter((doc) => isPdfEntry({ attributes: { displayName: doc.name, extension: doc.extension, mimeType: doc.mimeType } }, filePattern));
    const latestDoc = pdfDocs.sort((a, b) => String(b.updatedAt || b.createdAt || '').localeCompare(String(a.updatedAt || a.createdAt || '')))[0];
    if (!latestDoc?.id) {
      return res.status(404).json({ found: false, message: 'Relatório PDF não encontrado na pasta configurada.' });
    }
    const latestVersion = await getItemLatestVersion(req, projectId, latestDoc.id);
    const signedDownload = await callApsApi(req, `/oss/v2/buckets/${encodeURIComponent(parseOssStorageId(latestVersion?.relationships?.storage?.data?.id)?.bucket || '')}/objects/${encodeURIComponent(parseOssStorageId(latestVersion?.relationships?.storage?.data?.id)?.objectKey || '')}/signeds3download?minutesExpiration=10`);
    return res.json({
      found: true,
      file: {
        id: latestDoc.id,
        name: latestDoc.name,
        version: latestDoc.version || latestVersion?.attributes?.versionNumber || '',
        updatedAt: latestDoc.updatedAt || latestVersion?.attributes?.lastModifiedTime || '',
        folderPath: latestDoc.folderPath,
        downloadUrl: signedDownload.url || signedDownload.signedUrl || ''
      }
    });
  } catch (error) {
    next(error);
  }
});

app.post('/api/projects/:projectId/issues/relationships', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const issueIds = Array.isArray(req.body?.issueIds) ? req.body.issueIds.map((id) => String(id || '').trim()).filter(Boolean) : [];
    const results = {};

    await Promise.all(
      issueIds.map(async (issueId) => {
        const encodedIssueId = encodeURIComponent(issueId);
        const relationshipContainers = [...new Set([projectId, req.params.projectId].filter(Boolean))];
        const buildRelationshipSearch = (containerId, domain) => {
          const params = new URLSearchParams({
            domain,
            type: 'issue',
            id: issueId,
            withDomain: domain,
            withType: 'issue'
          });
          return `/bim360/relationship/v2/containers/${encodeURIComponent(containerId)}/relationships:search?${params.toString()}`;
        };
        const attempts = [
          ...relationshipContainers.flatMap((containerId) => [
            buildRelationshipSearch(containerId, 'autodesk-bim360-issue'),
            buildRelationshipSearch(containerId, 'autodesk-bim360-issues')
          ]),
          `/construction/issues/v1/projects/${projectId}/issues/${encodedIssueId}/relationships`,
          `/construction/issues/v1/projects/${projectId}/issues/${encodedIssueId}/references`
        ];

        for (const endpoint of attempts) {
          try {
            const payload = await callApsApi(req, endpoint);
            results[issueId] = payload?.relationships || payload?.results || payload?.data || payload || [];
            return;
          } catch (error) {
            // tenta o próximo endpoint
          }
        }
        results[issueId] = [];
      })
    );

    res.json({ results });
  } catch (error) {
    next(error);
  }
});

app.get('/api/projects/:projectId/issue-types', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const result = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-types?include=subtypes&limit=200`);
    res.json(listIssueTypes(result));
  } catch (error) {
    next(error);
  }
});

app.get('/api/projects/:projectId/issue-attribute-definitions', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const [result, mappingsResult] = await Promise.all([
      callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions?limit=200`),
      callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-mappings?limit=200`).catch(() => ({
        results: []
      }))
    ]);
    const mappingsByDefinitionId = mapAttributeMappings(mappingsResult);

    const definitions = (result.results || result.data || [])
      .filter(isActiveConfiguration)
      .map((definition) => {
        const id = definition.id || definition.attributeDefinitionId;

        return {
          id,
          name: definition.title || definition.name || definition.displayName || id,
          type: definition.dataType || definition.type || definition.metadata?.dataType || null,
          required: Boolean(definition.required || definition.isRequired),
          options: getDefinitionOptions(definition),
          appliesTo: mappingsByDefinitionId.get(id) || []
        };
      })
      .filter((definition) => definition.id);

    res.json(definitions);
  } catch (error) {
    next(error);
  }
});

app.get('/api/projects/:projectId/users', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const result = await callApsApi(req, `/construction/admin/v1/projects/${projectId}/users?limit=200`);
    res.json(listProjectUsers(result));
  } catch (error) {
    next(error);
  }
});

async function getIssueCreationContext(req, projectId) {
  const [issuesResult, issueTypesResult, definitionsResult, mappingsResult, usersResult] = await Promise.all([
    callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues?limit=200`).catch(() => ({ results: [] })),
    callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-types?include=subtypes&limit=200`),
    callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions?limit=200`).catch(() => ({
      results: []
    })),
    callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-mappings?limit=200`).catch(() => ({
      results: []
    })),
    callApsApi(req, `/construction/admin/v1/projects/${projectId}/users?limit=200`).catch(() => ({ results: [] }))
  ]);

  const mappingsByDefinitionId = mapAttributeMappings(mappingsResult);
  const fieldDefinitions = (definitionsResult.results || definitionsResult.data || [])
    .filter(isActiveConfiguration)
    .map((definition) => {
      const id = definition.id || definition.attributeDefinitionId;
      return {
        id,
        name: definition.title || definition.name || definition.displayName || id,
        type: definition.dataType || definition.type || definition.metadata?.dataType || null,
        required: Boolean(definition.required || definition.isRequired),
        options: getDefinitionOptions(definition),
        appliesTo: mappingsByDefinitionId.get(id) || []
      };
    })
    .filter((definition) => definition.id);

  return {
    issueTypes: listIssueTypes(issueTypesResult),
    fieldDefinitions,
    users: listProjectUsers(usersResult),
    existingIssues: (issuesResult.results || issuesResult.data || []).map((issue) => ({
      id: issue.id,
      title: readIssueField(issue, ['title'], ''),
      raw: issue
    }))
  };
}

function findProjectUser(users, value) {
  const normalizedValue = normalizeSpreadsheetKey(value);
  if (!normalizedValue) return null;

  return users.find((user) =>
    [user.id, user.name, user.email]
      .filter(Boolean)
      .some((userValue) => normalizeSpreadsheetKey(userValue) === normalizedValue || normalizeSpreadsheetKey(userValue).includes(normalizedValue))
  );
}

function resolveIssueType(issueTypes, manualRule, rowValues) {
  const candidates = [rowValues.issueType, manualRule?.type, manualRule?.category, rowValues.category].filter(Boolean);

  for (const candidate of candidates) {
    const exact = issueTypes.find((issueType) => normalizeSpreadsheetKey(issueType.title) === normalizeSpreadsheetKey(candidate));
    if (exact) return exact;
  }

  for (const candidate of candidates) {
    const found = findIssueTypeByName(issueTypes, candidate);
    if (found) return found;
  }

  return null;
}

function issueTypeAcceptsField(field, issueType) {
  if (!field.appliesTo?.length || !issueType) return true;

  return field.appliesTo.some((mapping) => {
    const mappingType = normalizeSpreadsheetKey(mapping.mappedItemType);
    const mappedId = mapping.mappedItemId;
    return (
      (mappedId === issueType.id && (mappingType.includes('type') || mappingType.includes('subtype'))) ||
      (issueType.typeId && mappedId === issueType.typeId && mappingType.includes('type'))
    );
  });
}

function findMappedColumnValue(values, columnMapping, targetKey, aliases) {
  const mappedColumn = columnMapping?.[targetKey];
  if (mappedColumn && Object.prototype.hasOwnProperty.call(values, mappedColumn)) {
    return values[mappedColumn];
  }
  return findColumnValue(values, aliases);
}

const eapFieldAliases = {
  eapCode: ['codigo eap', 'código eap', 'item eap', 'eap', 'codigo', 'código', 'codigo do issue', 'código do issue'],
  level: ['nivel', 'nível'],
  activity: ['atividade', 'servico', 'serviço', 'item', 'tarefa'],
  description: ['descricao', 'descrição', 'observacoes', 'observações', 'descricao complementar', 'descrição complementar'],
  discipline: ['disciplina', 'disciplinas envolvidas'],
  responsibleArea: ['area responsavel', 'área responsável', 'responsavel area', 'responsável área'],
  phase: ['fase'],
  gedDeliverable: ['entregavel ged', 'entregável ged', 'ged'],
  awpApplicable: ['aplicavel awp', 'aplicável awp', 'awp'],
  relatedPackage: ['pacote awp', 'pacote cwp', 'pacote iwp', 'pacote', 'awp', 'cwp', 'iwp', 'pacote marco relacionado'],
  relatedDocument: ['documento modelo relacionado', 'documento / modelo relacionado', 'documento relacionado', 'modelo relacionado'],
  measurementMilestone: ['marco de medicao', 'marco de medição', 'marco medicao', 'marco medição'],
  demandOrigin: ['origem da demanda', 'origem do comentario', 'origem do comentário', 'origem'],
  treatmentChannel: ['canal de tratativa', 'canal', 'tratativa'],
  requester: ['solicitante', 'requerente'],
  priority: ['prioridade'],
  impactExpected: ['nivel de impacto', 'nível de impacto', 'impacto previsto', 'impacto previsto:', 'impacto'],
  issueType: ['tipo de issue', 'tipo issue', 'tipo'],
  category: ['categoria do issue', 'categoria'],
  title: ['nome do issue', 'nome da issue', 'issue', 'titulo', 'título', 'assunto'],
  status: ['status inicial', 'status'],
  assignee: ['responsavel', 'responsável', 'responsavel oficial', 'responsável oficial', 'atribuido', 'atribuído'],
  dueDate: ['data prevista g5', 'data prevista', 'prazo de resposta', 'prazo', 'data limite', 'vencimento'],
  impactSchedule: ['impacto em prazo', 'impacto no cronograma', 'impacto no cronograma:', 'impacto prazo', 'impacto cronograma'],
  impactScope: ['impacto no escopo', 'impacto em escopo', 'impacto escopo'],
  impactMeasurement: ['impacto em medicao', 'impacto em medição', 'impacto na medição', 'impacto medicao', 'impacto medição'],
  observations: ['observacoes', 'observações', 'descricao complementar', 'descrição complementar', 'comentarios', 'comentários']
};

const eapCustomFieldAliases = {
  EAP: eapFieldAliases.eapCode,
  'Codigo do Issue': ['codigo do issue', 'código do issue', ...eapFieldAliases.eapCode],
  Nivel: eapFieldAliases.level,
  Atividade: eapFieldAliases.activity,
  Disciplina: eapFieldAliases.discipline,
  Fase: eapFieldAliases.phase,
  'Entregavel GED': eapFieldAliases.gedDeliverable,
  'Aplicavel AWP': eapFieldAliases.awpApplicable,
  'Origem da demanda': eapFieldAliases.demandOrigin,
  'Canal de tratativa': eapFieldAliases.treatmentChannel,
  Solicitante: eapFieldAliases.requester,
  'Responsavel oficial': eapFieldAliases.assignee,
  'Area responsavel': eapFieldAliases.responsibleArea,
  'Disciplinas envolvidas': eapFieldAliases.discipline,
  'Prazo de resposta': eapFieldAliases.dueDate,
  'Data Prevista G5': eapFieldAliases.dueDate,
  Prioridade: eapFieldAliases.priority,
  'Nível de impacto': eapFieldAliases.impactExpected,
  'Impacto em Prazo': eapFieldAliases.impactSchedule,
  'Impacto no Escopo': eapFieldAliases.impactScope,
  'Impacto em Medicao': eapFieldAliases.impactMeasurement,
  'Documento / Modelo relacionado': eapFieldAliases.relatedDocument,
  'Pacote / Marco relacionado': eapFieldAliases.relatedPackage,
  'Marco de Medicao': eapFieldAliases.measurementMilestone,
  Observacoes: eapFieldAliases.observations
};

function readEapRowValuesV2(row, columnMapping = {}) {
  const values = row.values;
  const eapCode = normalizeCellValue(
    findColumnValue(values, ['codigo eap', 'código eap', 'item eap', 'eap', 'codigo', 'código'])
  );
  const description = normalizeCellValue(findColumnValue(values, ['descricao', 'descrição', 'observacoes', 'observações']));
  const title = normalizeCellValue(findColumnValue(values, ['nome do issue', 'nome da issue', 'issue', 'titulo', 'título']));
  const dueDate = parseSpreadsheetDate(findColumnValue(values, ['data prevista g5', 'data prevista', 'prazo de resposta', 'prazo']));

  return {
    rowNumber: row.rowNumber,
    eapCode,
    description,
    discipline: normalizeCellValue(findColumnValue(values, ['disciplina'])),
    phase: normalizeCellValue(findColumnValue(values, ['fase'])),
    packageCode: normalizeCellValue(findColumnValue(values, ['pacote awp', 'pacote cwp', 'pacote iwp', 'awp', 'cwp', 'iwp'])),
    issueType: normalizeCellValue(findColumnValue(values, ['tipo de issue', 'tipo issue', 'tipo'])),
    category: normalizeCellValue(findColumnValue(values, ['categoria do issue', 'categoria'])),
    title,
    status: normalizeCellValue(findColumnValue(values, ['status inicial', 'status'])),
    assignee: normalizeCellValue(findColumnValue(values, ['responsavel', 'responsável', 'responsavel oficial', 'responsável oficial'])),
    dueDate,
    impactSchedule: normalizeCellValue(findColumnValue(values, ['impacto em prazo', 'impacto no cronograma', 'impacto prazo'])),
    impactScope: normalizeCellValue(findColumnValue(values, ['impacto no escopo', 'impacto em escopo', 'impacto escopo'])),
    impactMeasurement: normalizeCellValue(findColumnValue(values, ['impacto em medicao', 'impacto em medição', 'impacto na medição'])),
    observations: normalizeCellValue(findColumnValue(values, ['observacoes', 'observações', 'descricao complementar', 'descrição complementar'])),
    raw: values
  };
}

function buildEapIssuePreviewRows(workbookData, context = {}) {
  const issueTypes = Array.isArray(context.issueTypes) ? context.issueTypes : [];
  const fieldDefinitions = Array.isArray(context.fieldDefinitions) ? context.fieldDefinitions : [];
  const users = Array.isArray(context.users) ? context.users : [];
  const existingIssues = Array.isArray(context.existingIssues) ? context.existingIssues : [];
  const previewOnly = Boolean(context.previewOnly);
  const existingTitleKeys = new Set(existingIssues.map((issue) => normalizeSpreadsheetKey(issue.title)));
  const existingComboKeys = new Set(existingIssues.map((issue) => makeNormalizedIssueKey(getIssueCodeFromText(issue.title), issue.title)));
  const localTitleRows = new Map();
  const localComboRows = new Map();

  return workbookData.rows
    .map((row, index) => {
      const rowValues = readEapRowValues(row);
      const manualRule = pickManualIssueRule(rowValues);
      const title = buildIssueTitleFromEap(rowValues, manualRule);
      const issueType = resolveIssueType(issueTypes, null, { issueType: rowValues.issueType, category: rowValues.category });
      const user = findProjectUser(users, rowValues.assignee);
      const warnings = [];
      const errors = [];

      if (!rowValues.title && !rowValues.category && !rowValues.issueType) {
        return null;
      }

      if (!rowValues.eapCode) warnings.push('Linha sem código/item EAP.');
      if (!rowValues.description) warnings.push('Linha sem descrição.');
      if (!manualRule) warnings.push('Tipo/categoria não reconhecido nos manuais; usando dados da planilha quando possível.');
      if (rowValues.issueType && manualRule && normalizeSpreadsheetKey(rowValues.issueType) !== normalizeSpreadsheetKey(manualRule.type)) {
        warnings.push(`Manual priorizado: tipo "${manualRule.type}" aplicado no lugar de "${rowValues.issueType}".`);
      }
      if (!title) errors.push('Não foi possível definir o nome do Issue.');
      if (!issueType && !previewOnly) errors.push(`Tipo de Issue não encontrado no ACC para "${manualRule?.type || rowValues.issueType || rowValues.category}".`);
      if (previewOnly) warnings.push('Prévia inicial: a conferência final com tipos, usuários e duplicidades do ACC será feita antes da criação.');
      if (rowValues.assignee && !user) warnings.push(`Responsável "${rowValues.assignee}" não localizado no projeto.`);

      const applicableFields = fieldDefinitions.filter((field) => issueTypeAcceptsField(field, issueType));
      const customAttributes = buildCustomAttributesFromEap(rowValues, rowValues.raw, applicableFields);
      const duplicateKey = makeNormalizedIssueKey(getIssueCodeFromText(title), title);
      const duplicate = existingKeys.has(duplicateKey);
      if (duplicate) warnings.push('Já existe Issue com mesmo código/nome no projeto.');

      const payload = {
        title,
        description: [rowValues.description, rowValues.observations].filter(Boolean).join('\n\n'),
        dueDate: rowValues.dueDate,
        status: rowValues.status || 'open',
        assignedTo: user?.id || '',
        issueTypeId: issueType?.kind === 'type' ? issueType.id : '',
        issueSubtypeId: issueType?.kind === 'subtype' ? issueType.id : '',
        issueTypeName: issueType?.title || manualRule?.type || rowValues.issueType,
        customAttributes: customAttributes.map((field) => ({
          attributeDefinitionId: field.attributeDefinitionId,
          value: field.value
        }))
      };

      const validation = errors.length ? 'erro' : duplicate ? 'duplicado' : warnings.length ? 'incompleto' : 'pronto';

      return {
        id: `${row.rowNumber}-${index}`,
        line: row.rowNumber,
        eapCode: '',
        description: rowValues.description,
        discipline: rowValues.discipline,
        packageCode: rowValues.packageCode,
        title,
        issueType: issueType?.title || manualRule?.type || rowValues.issueType,
        category: issueType?.category || manualRule?.category || rowValues.category,
        dueDate: rowValues.dueDate,
        startDate: rowValues.startDate,
        status: payload.status,
        assignee: user?.name || rowValues.assignee,
        validation,
        errors,
        warnings,
        duplicate,
        payload
      };
    })
    .filter(Boolean);
}

function eapGetColumnLabelFromKey(key) {
  return String(key || '').replace(/__\d+$/, '').trim();
}

function eapNormalizeFieldValue(value) {
  const normalized = normalizeSpreadsheetKey(value);
  if (!normalized) return '';
  if (['sim', 'yes', 'true', '1'].includes(normalized)) return 'Sim';
  if (['nao', 'no', 'false', '0'].includes(normalized)) return 'Não';
  if (normalized === 'parcial') return 'Parcial';
  if (normalized === 'medio' || normalized === 'media') return 'Médio';
  if (normalized === 'alto' || normalized === 'alta') return 'Alto';
  if (normalized === 'baixo' || normalized === 'baixa') return 'Baixo';
  return normalizeCellValue(value);
}

function eapFindFieldDefinition(fieldDefinitions, fieldName, aliases = []) {
  const searchNames = [fieldName, ...aliases].filter(Boolean).map(normalizeSpreadsheetKey);
  return fieldDefinitions.find((field) => {
    const normalizedDefinition = normalizeSpreadsheetKey(field.name);
    return searchNames.some(
      (name) => normalizedDefinition === name || normalizedDefinition.includes(name) || name.includes(normalizedDefinition)
    );
  });
}

function eapCoerceCustomAttributeValue(definition, rawValue) {
  const value = normalizeCellValue(rawValue);
  if (!definition?.id || !value) return { value: '', displayValue: '', warning: '' };

  const type = normalizeSpreadsheetKey(definition.type || '');
  const options = Array.isArray(definition.options) ? definition.options : [];

  if (options.length > 0 || type.includes('list') || type.includes('dropdown') || type.includes('select')) {
    const normalizedValue = eapNormalizeFieldValue(value);
    const option = options.find(
      (candidate) =>
        fieldValueMatchesOption(candidate, value) ||
        fieldValueMatchesOption(candidate, normalizedValue) ||
        normalizeSpreadsheetKey(candidate.label || candidate.name || candidate.value) === normalizeSpreadsheetKey(normalizedValue)
    );

    if (!option) {
      return {
        value: '',
        displayValue: value,
        warning: `Valor "${value}" nao existe nas opcoes do campo "${definition.name}".`
      };
    }

    return {
      value: option.id || option.value || option.label,
      displayValue: option.label || option.name || option.value || value,
      warning: ''
    };
  }

  if (type.includes('date')) {
    const dateValue = parseSpreadsheetDate(value);
    return dateValue
      ? { value: dateValue, displayValue: dateValue, warning: '' }
      : { value: '', displayValue: value, warning: `Data invalida no campo "${definition.name}".` };
  }

  if (type.includes('number') || type.includes('integer') || type.includes('double')) {
    const numericValue = Number(String(value).replace(/\./g, '').replace(',', '.'));
    return Number.isFinite(numericValue)
      ? { value: numericValue, displayValue: String(numericValue), warning: '' }
      : { value: '', displayValue: value, warning: `Numero invalido no campo "${definition.name}".` };
  }

  if (type.includes('bool')) {
    const normalized = normalizeSpreadsheetKey(value);
    if (['sim', 'yes', 'true', '1'].includes(normalized)) return { value: true, displayValue: 'Sim', warning: '' };
    if (['nao', 'no', 'false', '0'].includes(normalized)) return { value: false, displayValue: 'Não', warning: '' };
    return { value: '', displayValue: value, warning: `Valor sim/nao invalido no campo "${definition.name}".` };
  }

  return { value: eapNormalizeFieldValue(value), displayValue: eapNormalizeFieldValue(value), warning: '' };
}

function eapBuildManualCustomFieldValues(rowValues) {
  return {
    EAP: rowValues.eapCode,
    'Codigo do Issue': rowValues.eapCode,
    Nivel: rowValues.level,
    Atividade: rowValues.activity || rowValues.description,
    Disciplina: rowValues.discipline,
    Fase: rowValues.phase,
    'Entregavel GED': rowValues.gedDeliverable,
    'Aplicavel AWP': rowValues.awpApplicable,
    'Origem da demanda': rowValues.demandOrigin,
    'Canal de tratativa': rowValues.treatmentChannel,
    Solicitante: rowValues.requester,
    'Responsavel oficial': rowValues.assignee,
    'Area responsavel': rowValues.responsibleArea,
    'Disciplinas envolvidas': rowValues.discipline,
    'Prazo de resposta': rowValues.dueDate,
    'Data Prevista G5': rowValues.dueDate,
    Prioridade: rowValues.priority,
    'Nível de impacto': rowValues.impactExpected,
    'Impacto em Prazo': mapImpactValue(rowValues.impactSchedule),
    'Impacto no Escopo': mapImpactValue(rowValues.impactScope),
    'Impacto em Medicao': mapImpactValue(rowValues.impactMeasurement),
    'Documento / Modelo relacionado': rowValues.relatedDocument,
    'Pacote / Marco relacionado': rowValues.relatedPackage || rowValues.packageCode,
    'Marco de Medicao': rowValues.measurementMilestone,
    Observacoes: rowValues.observations
  };
}

function buildCustomAttributesFromEap(rowValues, row, fieldDefinitions) {
  const attributes = [];
  const filled = [];
  const missing = [];
  const warnings = [];
  const fieldValues = eapBuildManualCustomFieldValues(rowValues);

  for (const [key, value] of Object.entries(row || {})) {
    const cleanKey = eapGetColumnLabelFromKey(key);
    if (!cleanKey || !normalizeCellValue(value)) continue;
    const nativeOnlyAliases = [
      ...eapFieldAliases.title,
      ...eapFieldAliases.issueType,
      ...eapFieldAliases.category,
      ...eapFieldAliases.status
    ];
    if (nativeOnlyAliases.some((alias) => normalizeSpreadsheetKey(cleanKey) === normalizeSpreadsheetKey(alias))) continue;
    fieldValues[cleanKey] = normalizeCellValue(value);
  }

  for (const [fieldName, rawValue] of Object.entries(fieldValues)) {
    const value = normalizeCellValue(rawValue);
    if (!value) continue;

    const aliases = eapCustomFieldAliases[fieldName] || [];
    const definition = eapFindFieldDefinition(fieldDefinitions, fieldName, aliases);

    if (!definition?.id) {
      missing.push(fieldName);
      continue;
    }

    const coerced = eapCoerceCustomAttributeValue(definition, value);
    if (coerced.warning) {
      warnings.push(definition.required ? `${coerced.warning} Campo obrigatorio.` : coerced.warning);
      continue;
    }

    if (coerced.value === '' || coerced.value === undefined || coerced.value === null) continue;

    attributes.push({
      attributeDefinitionId: definition.id,
      name: definition.name,
      value: coerced.value,
      displayValue: coerced.displayValue
    });
    filled.push(definition.name);
  }

  return { attributes, filled: [...new Set(filled)], missing: [...new Set(missing)], warnings };
}

function parseEapWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = workbook.SheetNames.find((name) => normalizeSpreadsheetKey(name) === normalizeSpreadsheetKey('ACS Build'));

  if (!sheetName) {
    const error = new Error('A planilha enviada nao possui uma aba chamada ACS Build.');
    error.status = 400;
    throw error;
  }

  const tableRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false });
  const requiredHeaders = ['title', 'status', 'category', 'type'];
  const headerRowIndex = 0;
  const friendlyHeaderRowIndex = 1;
  const reminderRowIndex = 2;
  const purpleHeader = tableRows[headerRowIndex] || [];
  const greenHeader = tableRows[friendlyHeaderRowIndex] || [];
  const maxColumns = Math.max(purpleHeader.length, greenHeader.length);
  const columns = Array.from({ length: maxColumns }).map((_, index) => {
    const issueField = String(purpleHeader[index] || '').trim();
    const friendlyField = String(greenHeader[index] || '').trim();
    const label = issueField || `Coluna ${index + 1}`;
    return {
      key: `${label}__${index}`,
      label,
      issueField,
      friendlyField,
      index
    };
  });

  const normalizedColumnNames = columns.map((column) => normalizeSpreadsheetKey(column.issueField || column.friendlyField || column.label));
  const missingRequired = requiredHeaders.filter((header) => !normalizedColumnNames.includes(normalizeSpreadsheetKey(header)));
  if (missingRequired.length) {
    const error = new Error(`Colunas obrigatorias ausentes na aba ACS Build: ${missingRequired.join(', ')}.`);
    error.status = 400;
    throw error;
  }

  const firstDataRowIndex = 3;
  const rows = tableRows.slice(firstDataRowIndex).map((row, rowIndex) => ({
    rowNumber: firstDataRowIndex + rowIndex + 1,
    values: columns.reduce((record, column) => {
      record[column.key] = row[column.index] ?? '';
      return record;
    }, {})
  })).filter((row) => {
    const values = Object.values(row.values || {}).map((value) => normalizeCellValue(value));
    if (!values.some(Boolean)) return false;
    if (normalizeSpreadsheetKey(values.join(' | ')).includes('reminders:')) return false;
    const category = normalizeCellValue(findColumnValue(row.values, ['category']));
    const type = normalizeCellValue(findColumnValue(row.values, ['type']));
    return Boolean(category || type);
  });

  const dropdownOptionsByColumn = columns.reduce((acc, column) => {
    const instruction = String(greenHeader[column.index] || '').trim();
    if (!instruction.includes(';')) return acc;
    const options = instruction.split(';').map((item) => item.trim()).filter(Boolean);
    if (options.length >= 2) acc[column.label] = options;
    return acc;
  }, {});

  return {
    sheetName,
    headerRowIndex,
    friendlyHeaderRowIndex,
    reminderRowIndex,
    columns,
    rows,
    dropdownOptionsByColumn
  };
}

function readEapRowValues(row, columnMapping = {}) {
  const values = row.values || {};
  const get = (...aliases) => normalizeCellValue(findColumnValue(values, aliases));

  return {
    rowNumber: row.rowNumber,
    title: get('title'),
    status: get('status'),
    category: get('category'),
    issueType: get('type'),
    description: get('description'),
    location: get('location'),
    assignee: get('assigned to'),
    assigneeType: get('assignee type'),
    dueDate: parseSpreadsheetDate(findColumnValue(values, ['due date'])),
    startDate: parseSpreadsheetDate(findColumnValue(values, ['start date'])),
    rootCause: get('root cause'),
    eapCode: get('codigo eap', 'código eap'),
    raw: values
  };
}

function buildIssueTitleFromEapV2(rowValues, manualRule) {
  if (rowValues.title) return rowValues.title;

  const typePrefix = manualRule?.code || manualRule?.type || rowValues.issueType || 'ISSUE';
  const code = rowValues.eapCode || rowValues.packageCode || '';
  const origin = rowValues.demandOrigin || rowValues.discipline || rowValues.category || 'Origem';
  const responsible = rowValues.responsibleArea || rowValues.assignee || 'Responsavel';
  const description = rowValues.activity || rowValues.description || rowValues.observations || 'Item da EAP';
  const compactDescription = description.length > 70 ? `${description.slice(0, 67)}...` : description;
  return `${typePrefix}${code ? `-${code}` : ''} | ${origin} x ${responsible} | ${compactDescription}`;
}

function buildEapIssueDescription(rowValues) {
  const lines = [
    ['Codigo EAP', rowValues.eapCode],
    ['Atividade', rowValues.activity || rowValues.description],
    ['Disciplina', rowValues.discipline],
    ['Area responsavel', rowValues.responsibleArea],
    ['Fase', rowValues.phase],
    ['Entregavel GED', rowValues.gedDeliverable],
    ['Aplicavel AWP', rowValues.awpApplicable],
    ['Impacto em prazo', mapImpactValue(rowValues.impactSchedule)],
    ['Impacto no escopo', mapImpactValue(rowValues.impactScope)],
    ['Impacto em medicao', mapImpactValue(rowValues.impactMeasurement)],
    ['Pacote / Marco relacionado', rowValues.relatedPackage || rowValues.packageCode],
    ['Documento / Modelo relacionado', rowValues.relatedDocument],
    ['Observacoes', rowValues.observations]
  ];

  return lines
    .filter(([, value]) => normalizeCellValue(value))
    .map(([label, value]) => `${label}: ${normalizeCellValue(value)}`)
    .join('\n');
}

function buildEapIssuePreviewRowsV2(workbookData, context = {}) {
  const issueTypes = Array.isArray(context.issueTypes) ? context.issueTypes : [];
  const fieldDefinitions = Array.isArray(context.fieldDefinitions) ? context.fieldDefinitions : [];
  const users = Array.isArray(context.users) ? context.users : [];
  const existingIssues = Array.isArray(context.existingIssues) ? context.existingIssues : [];
  const columnMapping = context.columnMapping || {};
  const previewOnly = Boolean(context.previewOnly);
  const existingTitleKeys = new Set(existingIssues.map((issue) => normalizeSpreadsheetKey(issue.title)));
  const existingComboKeys = new Set(existingIssues.map((issue) => makeNormalizedIssueKey(getIssueCodeFromText(issue.title), issue.title)));
  const localTitleRows = new Map();
  const localComboRows = new Map();

  return workbookData.rows
    .map((row, index) => {
      const rowValues = readEapRowValues(row, columnMapping);
      const manualRule = pickManualIssueRule(rowValues);
      const title = rowValues.title;
      const issueType = resolveIssueType(issueTypes, null, { issueType: rowValues.issueType, category: rowValues.category });
      const user = findProjectUser(users, rowValues.assignee);
      const warnings = [];
      const errors = [];

      if (!rowValues.title && !rowValues.category && !rowValues.issueType) {
        return null;
      }

      if (!rowValues.title) errors.push('Linha sem Title.');
      if (!rowValues.category) errors.push('Linha sem Category.');
      if (!rowValues.issueType) errors.push('Linha sem Type.');
      if (!rowValues.status) errors.push('Linha sem Status.');
      if (!title) errors.push('Nao foi possivel definir o nome do Issue.');
      if (!issueType && !previewOnly) errors.push(`Tipo de Issue nao encontrado no ACC para "${manualRule?.type || rowValues.issueType || rowValues.category}".`);
      if (previewOnly) warnings.push('Previa inicial: a conferencia final com tipos, usuarios, campos personalizados e duplicidades do ACC sera feita antes da criacao.');
      if (rowValues.assignee && !user && !previewOnly) warnings.push(`Responsavel "${rowValues.assignee}" nao localizado no projeto; nao sera atribuido automaticamente.`);

      const applicableFields = fieldDefinitions.filter((field) => issueTypeAcceptsField(field, issueType));
      const customAttributeResult = buildCustomAttributesFromEap(rowValues, rowValues.raw, applicableFields) || {};
      const customAttributeWarnings = Array.isArray(customAttributeResult.warnings) ? customAttributeResult.warnings : [];
      const customAttributeMissing = Array.isArray(customAttributeResult.missing) ? customAttributeResult.missing : [];
      const customAttributes = Array.isArray(customAttributeResult.attributes) ? customAttributeResult.attributes : [];
      warnings.push(...customAttributeWarnings);
      if (!previewOnly && customAttributeMissing.length) {
        warnings.push(`Campos da planilha/manual sem correspondente no ACC: ${customAttributeMissing.slice(0, 8).join(', ')}.`);
      }

      const normalizedTitle = normalizeSpreadsheetKey(title);
      const comboKey = rowValues.eapCode ? makeNormalizedIssueKey(rowValues.eapCode, title) : '';
      let duplicate = false;
      if (comboKey && (existingComboKeys.has(comboKey) || localComboRows.has(comboKey))) {
        duplicate = true;
        warnings.push(
          `Duplicado por combinação Title + Código EAP${localComboRows.has(comboKey) ? ` (linha ${localComboRows.get(comboKey)}).` : ' já existente no ACC.'}`
        );
      } else if (!comboKey && normalizedTitle && (existingTitleKeys.has(normalizedTitle) || localTitleRows.has(normalizedTitle))) {
        duplicate = true;
        warnings.push(`Duplicado por Title${localTitleRows.has(normalizedTitle) ? ` (linha ${localTitleRows.get(normalizedTitle)}).` : ' já existente no ACC.'}`);
      }
      if (normalizedTitle) localTitleRows.set(normalizedTitle, row.rowNumber);
      if (comboKey) localComboRows.set(comboKey, row.rowNumber);

      const payload = {
        title,
        description: rowValues.description || '',
        dueDate: rowValues.dueDate,
        status: rowValues.status || 'open',
        assignedTo: user?.id || '',
        issueTypeId: issueType?.kind === 'type' ? issueType.id : '',
        issueSubtypeId: issueType?.kind === 'subtype' ? issueType.id : '',
        issueTypeName: issueType?.title || manualRule?.type || rowValues.issueType,
        customAttributes: customAttributes.map((field) => ({
          attributeDefinitionId: field.attributeDefinitionId,
          value: field.value
        }))
      };

      const validation = errors.length ? 'erro' : duplicate ? 'duplicado' : warnings.length ? 'incompleto' : 'pronto';

      return {
        id: `${row.rowNumber}-${index}`,
        line: row.rowNumber,
        eapCode: rowValues.eapCode || '',
        description: rowValues.description,
        title,
        issueType: issueType?.title || manualRule?.type || rowValues.issueType,
        category: issueType?.category || manualRule?.category || rowValues.category,
        dueDate: rowValues.dueDate,
        startDate: rowValues.startDate,
        status: payload.status,
        assignee: user?.name || rowValues.assignee,
        validation,
        errors,
        warnings,
        duplicate,
        customFieldsFilled: customAttributeResult.filled,
        customFieldsMissing: customAttributeResult.missing,
        nativeFieldsFilled: Object.entries(payload).filter(([, value]) => value && (!Array.isArray(value) || value.length)).map(([key]) => key),
        sourceValues: row.values,
        payload
      };
    })
    .filter(Boolean);
}

const imFieldAliases = {
  action: ['acao', 'ação'],
  issueId: ['issue_id', 'id issue', 'id do issue', 'issue acc', 'id acc'],
  importId: ['identificador_importacao', 'identificador importacao', 'identificador importação', 'import id'],
  project: ['projeto', 'project'],
  title: ['titulo', 'título', 'title', 'nome do issue', 'nome issue'],
  description: ['descricao', 'descrição', 'description'],
  category: ['categoria', 'categoria do issue', 'category'],
  issueType: ['tipo', 'tipo de issue', 'issue type', 'type'],
  subtype: ['subtipo', 'subtipo de issue', 'issue subtype'],
  status: ['status', 'status inicial'],
  assignee: ['responsavel', 'responsável', 'atribuido a', 'atribuído a', 'responsavel oficial'],
  dueDate: ['data_vencimento', 'data vencimento', 'vencimento', 'due date', 'data prevista g5', 'prazo'],
  eapCode: ['codigo_eap', 'código eap', 'codigo eap', 'eap', 'código do issue', 'codigo do issue'],
  itemEap: ['item_eap', 'item eap', 'atividade'],
  packageCode: ['pacote_trabalho', 'pacote trabalho', 'pacote awp', 'pacote cwp', 'pacote iwp'],
  discipline: ['disciplina', 'discipline'],
  phase: ['fase', 'phase'],
  milestone: ['marco_contratual', 'marco contratual', 'marco'],
  impactSchedule: ['impacto_prazo', 'impacto prazo', 'impacto em prazo', 'impacto no cronograma'],
  impactScope: ['impacto_escopo', 'impacto escopo', 'impacto no escopo'],
  impactMeasurement: ['impacto_medicao', 'impacto medição', 'impacto medicao', 'impacto na medição', 'impacto em medicao'],
  priority: ['prioridade', 'priority'],
  observations: ['observacoes', 'observações', 'observacao', 'observação'],
  location: ['localizacao', 'localização', 'location'],
  clearFields: ['limpar_campos', 'limpar campos']
};

const imCustomFieldMap = {
  'Código EAP': ['codigo_eap', 'codigo eap', 'código eap', 'eap'],
  'Item EAP': ['item_eap', 'item eap', 'atividade'],
  'Pacote de Trabalho': ['pacote_trabalho', 'pacote trabalho', 'pacote awp', 'pacote cwp', 'pacote iwp'],
  Disciplina: ['disciplina'],
  Fase: ['fase'],
  'Marco Contratual': ['marco_contratual', 'marco contratual', 'marco'],
  'Impacto no Prazo': ['impacto_prazo', 'impacto prazo', 'impacto em prazo', 'impacto no cronograma'],
  'Impacto no Escopo': ['impacto_escopo', 'impacto escopo', 'impacto no escopo'],
  'Impacto na Medição': ['impacto_medicao', 'impacto medição', 'impacto medicao', 'impacto em medicao'],
  Prioridade: ['prioridade'],
  Observações: ['observacoes', 'observações', 'observacao', 'observação']
};

function readImRowValues(row, columnMapping = {}) {
  const values = row.values || {};
  const get = (targetKey, aliases) => normalizeCellValue(findMappedColumnValue(values, columnMapping, targetKey, aliases));

  const title = get('title', imFieldAliases.title);
  const description = get('description', imFieldAliases.description) || get('itemEap', imFieldAliases.itemEap);
  const dueDate = parseSpreadsheetDate(findMappedColumnValue(values, columnMapping, 'dueDate', imFieldAliases.dueDate));

  return {
    rowNumber: row.rowNumber,
    action: normalizeSpreadsheetKey(get('action', imFieldAliases.action)),
    issueId: get('issueId', imFieldAliases.issueId),
    importId: get('importId', imFieldAliases.importId),
    project: get('project', imFieldAliases.project),
    title,
    description,
    category: get('category', imFieldAliases.category),
    issueType: get('issueType', imFieldAliases.issueType),
    subtype: get('subtype', imFieldAliases.subtype),
    status: get('status', imFieldAliases.status),
    assignee: get('assignee', imFieldAliases.assignee),
    dueDate,
    eapCode: get('eapCode', imFieldAliases.eapCode),
    itemEap: get('itemEap', imFieldAliases.itemEap),
    packageCode: get('packageCode', imFieldAliases.packageCode),
    discipline: get('discipline', imFieldAliases.discipline),
    phase: get('phase', imFieldAliases.phase),
    milestone: get('milestone', imFieldAliases.milestone),
    impactSchedule: get('impactSchedule', imFieldAliases.impactSchedule),
    impactScope: get('impactScope', imFieldAliases.impactScope),
    impactMeasurement: get('impactMeasurement', imFieldAliases.impactMeasurement),
    priority: get('priority', imFieldAliases.priority),
    observations: get('observations', imFieldAliases.observations),
    location: get('location', imFieldAliases.location),
    clearFields: get('clearFields', imFieldAliases.clearFields),
    raw: values
  };
}

function getImColumnLabel(key) {
  return eapGetColumnLabelFromKey(key).replace(/^cf[_\s-]*/i, '').trim();
}

function isImCustomColumn(key) {
  return /^cf[_\s-]*/i.test(eapGetColumnLabelFromKey(key));
}

function buildImIssueTitle(rowValues) {
  if (rowValues.title) return rowValues.title;
  const code = rowValues.eapCode || rowValues.itemEap || '';
  const type = rowValues.issueType || rowValues.subtype || 'Issue';
  const text = rowValues.description || rowValues.observations || 'Item sem descricao';
  const compact = text.length > 70 ? `${text.slice(0, 67)}...` : text;
  return [type, code, compact].filter(Boolean).join(' | ');
}

function buildImIssueDescription(rowValues) {
  const lines = [
    ['Origem', 'IM - Criar Issues'],
    ['Codigo EAP', rowValues.eapCode],
    ['Item EAP', rowValues.itemEap],
    ['Descricao', rowValues.description],
    ['Disciplina', rowValues.discipline],
    ['Pacote de trabalho', rowValues.packageCode],
    ['Fase', rowValues.phase],
    ['Marco contratual', rowValues.milestone],
    ['Impacto prazo', mapImpactValue(rowValues.impactSchedule)],
    ['Impacto escopo', mapImpactValue(rowValues.impactScope)],
    ['Impacto medicao', mapImpactValue(rowValues.impactMeasurement)],
    ['Prioridade', rowValues.priority],
    ['Observacoes', rowValues.observations]
  ];

  return lines
    .filter(([, value]) => normalizeCellValue(value))
    .map(([label, value]) => `${label}: ${normalizeCellValue(value)}`)
    .join('\n');
}

function resolveImIssueType(issueTypes, rowValues) {
  const normalizedCategory = normalizeSpreadsheetKey(rowValues.category);
  const normalizedSubtype = normalizeSpreadsheetKey(rowValues.subtype);
  const normalizedType = normalizeSpreadsheetKey(rowValues.issueType);

  if (normalizedSubtype) {
    const subtype = issueTypes.find((item) => item.kind === 'subtype' && normalizeSpreadsheetKey(item.title) === normalizedSubtype);
    if (subtype && (!normalizedCategory || normalizeSpreadsheetKey(subtype.category) === normalizedCategory)) return subtype;
  }

  if (normalizedType) {
    const exact = issueTypes.find((item) => normalizeSpreadsheetKey(item.title) === normalizedType);
    if (exact && (!normalizedCategory || normalizeSpreadsheetKey(exact.category) === normalizedCategory)) return exact;
  }

  if (normalizedCategory) {
    return issueTypes.find((item) => normalizeSpreadsheetKey(item.category || item.title) === normalizedCategory) || null;
  }

  return null;
}

function buildImCustomAttributes(rowValues, fieldDefinitions, issueType) {
  const fieldValues = {};
  const warnings = [];
  const missing = [];
  const applicableFields = fieldDefinitions.filter((field) => issueTypeAcceptsField(field, issueType));

  for (const [fieldName, aliases] of Object.entries(imCustomFieldMap)) {
    const rawValue = findMappedColumnValue(rowValues.raw, {}, fieldName, aliases);
    if (normalizeCellValue(rawValue)) fieldValues[fieldName] = rawValue;
  }

  for (const [key, value] of Object.entries(rowValues.raw || {})) {
    if (!normalizeCellValue(value) || !isImCustomColumn(key)) continue;
    fieldValues[getImColumnLabel(key)] = value;
  }

  const attributes = [];
  const filled = [];

  for (const [fieldName, rawValue] of Object.entries(fieldValues)) {
    const value = normalizeCellValue(rawValue);
    if (!value) continue;
    const definition = eapFindFieldDefinition(applicableFields, fieldName, imCustomFieldMap[fieldName] || []);
    if (!definition?.id) {
      missing.push(fieldName);
      continue;
    }

    const coerced = eapCoerceCustomAttributeValue(definition, value);
    if (coerced.warning) {
      warnings.push(coerced.warning);
      continue;
    }

    if (coerced.value === '' || coerced.value === undefined || coerced.value === null) continue;
    attributes.push({
      attributeDefinitionId: definition.id,
      name: definition.name,
      value: coerced.value,
      displayValue: coerced.displayValue
    });
    filled.push(definition.name);
  }

  for (const definition of applicableFields) {
    if (definition.required && !attributes.some((item) => item.attributeDefinitionId === definition.id)) {
      warnings.push(`Campo personalizado obrigatorio sem valor: ${definition.name}.`);
    }
  }

  return { attributes, filled: [...new Set(filled)], missing: [...new Set(missing)], warnings };
}

function getIssueCustomAttributeValue(issue, fieldName) {
  const normalizedField = normalizeSpreadsheetKey(fieldName);
  const attributes = issue.customAttributes || normalizeCustomAttributes(readIssueField(issue.raw || {}, ['customAttributes', 'custom_fields'], []));
  const match = attributes.find((attribute) => normalizeSpreadsheetKey(attribute.name || attribute.id) === normalizedField);
  return normalizeCellValue(match?.displayValue ?? match?.value);
}

function makeImTraceabilityKey(rowValues) {
  if (rowValues.issueId) return makeNormalizedIssueKey(rowValues.issueId);
  if (rowValues.eapCode && rowValues.title) return makeNormalizedIssueKey(rowValues.eapCode, rowValues.title);
  if (rowValues.eapCode && rowValues.itemEap) return makeNormalizedIssueKey(rowValues.eapCode, rowValues.itemEap);
  if (rowValues.importId) return makeNormalizedIssueKey(rowValues.importId);
  return '';
}

function mapExistingIssueForIm(issue, fieldDefinitions = []) {
  const definitionMap = new Map(fieldDefinitions.map((field) => [field.id, field]));
  return mapIssue(issue.raw || issue, definitionMap);
}

function findExistingIssueForIm(rowValues, existingIssues, fieldDefinitions = []) {
  const mappedIssues = existingIssues.map((issue) => mapExistingIssueForIm(issue, fieldDefinitions));
  if (rowValues.issueId) {
    return mappedIssues.find((issue) => makeNormalizedIssueKey(issue.id) === makeNormalizedIssueKey(rowValues.issueId)) || null;
  }

  if (rowValues.eapCode) {
    const byCustomField = mappedIssues.find(
      (issue) => normalizeSpreadsheetKey(getIssueCustomAttributeValue(issue, 'Código EAP')) === normalizeSpreadsheetKey(rowValues.eapCode)
    );
    if (byCustomField) return byCustomField;
  }

  const byCodeAndTitle = makeImTraceabilityKey(rowValues);
  if (byCodeAndTitle) {
    return mappedIssues.find((issue) => makeNormalizedIssueKey(rowValues.eapCode, issue.title) === byCodeAndTitle) || null;
  }

  return null;
}

function buildImIssuePreviewRows(workbookData, context = {}, mode = 'create') {
  const issueTypes = Array.isArray(context.issueTypes) ? context.issueTypes : [];
  const fieldDefinitions = Array.isArray(context.fieldDefinitions) ? context.fieldDefinitions : [];
  const users = Array.isArray(context.users) ? context.users : [];
  const existingIssues = Array.isArray(context.existingIssues) ? context.existingIssues : [];
  const previewOnly = Boolean(context.previewOnly);
  const rows = [];

  for (const [index, row] of workbookData.rows.entries()) {
    const rowValues = readImRowValues(row);
    const action = rowValues.action;
    const isUpdateMode = mode === 'update';
    const actionMatches = isUpdateMode
      ? action === 'atualizar' || Boolean(rowValues.issueId)
      : !action || action === 'criar';
    if (!actionMatches) continue;
    if (!rowValues.title && !rowValues.eapCode && !rowValues.description && !rowValues.itemEap) continue;

    const title = buildImIssueTitle(rowValues);
    const issueType = resolveImIssueType(issueTypes, rowValues);
    const user = findProjectUser(users, rowValues.assignee);
    const existingIssue = findExistingIssueForIm(rowValues, existingIssues, fieldDefinitions);
    const customAttributeResult = buildImCustomAttributes(rowValues, fieldDefinitions, issueType);
    const warnings = [...customAttributeResult.warnings];
    const errors = [];

    if (!title) errors.push('Titulo vazio.');
    if (!previewOnly && !issueType) errors.push(`Categoria/tipo nao encontrado no ACC: ${rowValues.category || rowValues.issueType || rowValues.subtype || 'sem informacao'}.`);
    if (rowValues.assignee && !user && !previewOnly) warnings.push(`Responsavel "${rowValues.assignee}" nao encontrado; campo nativo nao sera atribuido.`);
    if (!previewOnly && customAttributeResult.missing.length) warnings.push(`Campos personalizados nao encontrados no ACC: ${customAttributeResult.missing.join(', ')}.`);

    if (isUpdateMode && !existingIssue && !previewOnly) errors.push('Nao encontrei correspondencia segura para atualizar este issue.');
    if (!isUpdateMode && existingIssue && !previewOnly) warnings.push('Ja existe um issue com a mesma rastreabilidade; a linha sera tratada como duplicada.');
    if (previewOnly) warnings.push('Previa inicial: a validacao final com metadados do ACC acontece antes da execucao.');

    const payload = {
      title,
      description: rowValues.description ? buildImIssueDescription(rowValues) : '',
      dueDate: rowValues.dueDate,
      status: normalizeSpreadsheetKey(rowValues.status).includes('fech') || normalizeSpreadsheetKey(rowValues.status).includes('closed') ? 'closed' : rowValues.status || 'open',
      assignedTo: user?.id || '',
      issueTypeId: issueType?.kind === 'type' ? issueType.id : '',
      issueSubtypeId: issueType?.kind === 'subtype' ? issueType.id : '',
      issueTypeName: issueType?.title || rowValues.subtype || rowValues.issueType,
      customAttributes: customAttributeResult.attributes.map((field) => ({
        attributeDefinitionId: field.attributeDefinitionId,
        value: field.value
      }))
    };

    const validation = errors.length ? 'erro' : !isUpdateMode && existingIssue ? 'duplicado' : warnings.length ? 'alerta' : 'pronto';
    rows.push({
      id: `${mode}-${row.rowNumber}-${index}`,
      mode,
      line: row.rowNumber,
      action: isUpdateMode ? 'atualizar' : 'criar',
      eapCode: rowValues.eapCode,
      itemEap: rowValues.itemEap,
      description: rowValues.description,
      title,
      category: rowValues.category || issueType?.category || '',
      issueType: rowValues.subtype || rowValues.issueType || issueType?.title || '',
      status: payload.status,
      assignee: rowValues.assignee,
      dueDate: rowValues.dueDate,
      impacts: {
        schedule: mapImpactValue(rowValues.impactSchedule),
        scope: mapImpactValue(rowValues.impactScope),
        measurement: mapImpactValue(rowValues.impactMeasurement)
      },
      traceabilityKey: makeImTraceabilityKey(rowValues),
      existingIssueId: existingIssue?.id || '',
      existingIssueTitle: existingIssue?.title || '',
      validation,
      errors,
      warnings,
      nativeFieldsFilled: Object.entries(payload)
        .filter(([key, value]) => key !== 'customAttributes' && normalizeCellValue(value))
        .map(([key]) => key),
      customFieldsFilled: customAttributeResult.filled,
      customFieldsMissing: customAttributeResult.missing,
      payload
    });
  }

  return rows;
}

function cleanIssuePayload(payload, preserveEmpty = false) {
  const cleanPayload = {};
  for (const [key, value] of Object.entries(payload || {})) {
    if (key === 'issueTypeName') continue;
    if (key === 'customAttributes') {
      if (Array.isArray(value) && value.length > 0) cleanPayload.customAttributes = value;
      continue;
    }
    if (value === undefined || value === null || value === '') {
      if (preserveEmpty) cleanPayload[key] = null;
      continue;
    }
    cleanPayload[key] = value;
  }
  return cleanPayload;
}

function decodeBase64Workbook(fileBase64) {
  const base64 = String(fileBase64 || '').replace(/^data:.*?;base64,/, '');
  if (!base64) {
    const error = new Error('Envie um arquivo Excel para leitura.');
    error.status = 400;
    throw error;
  }
  return Buffer.from(base64, 'base64');
}

function normalizeUploadedWorkbookData(workbookData) {
  if (!workbookData || !Array.isArray(workbookData.rows)) {
    return null;
  }

  return {
    sheetName: workbookData.sheetName || 'ACS Build',
    headerRowIndex: Number(workbookData.headerRowIndex || 0),
    friendlyHeaderRowIndex: Number(workbookData.friendlyHeaderRowIndex ?? -1),
    columns: Array.isArray(workbookData.columns) ? workbookData.columns : [],
    dropdownOptionsByColumn: workbookData.dropdownOptionsByColumn && typeof workbookData.dropdownOptionsByColumn === 'object' ? workbookData.dropdownOptionsByColumn : {},
    rows: workbookData.rows
      .map((row, index) => ({
        rowNumber: Number(row.rowNumber || index + 2),
        values: row.values && typeof row.values === 'object' ? row.values : {}
      }))
      .filter((row) => Object.values(row.values).some((value) => normalizeCellValue(value)))
  };
}

function normalizeImHeaders(value) {
  return normalizeSpreadsheetKey(String(value || '').replace(/^cf_/i, '').trim());
}

function matchImIssue(existingIssues, item) {
  const itemEapCode = item.codigoEap || item.eapCode || '';
  const byId = item.issueId
    ? existingIssues.filter((issue) => normalizeSpreadsheetKey(issue.id) === normalizeSpreadsheetKey(item.issueId))
    : [];
  const byCode = !byId.length && itemEapCode
    ? existingIssues.filter((issue) => normalizeSpreadsheetKey(issue.displayId || '').includes(normalizeSpreadsheetKey(itemEapCode)))
    : [];
  const byTitle = !byId.length && !byCode.length && item.title
    ? existingIssues.filter((issue) => normalizeSpreadsheetKey(issue.title) === normalizeSpreadsheetKey(item.title))
    : [];
  const matches = byId.length ? byId : byCode.length ? byCode : byTitle;
  return {
    matches,
    found: matches.length === 1 ? matches[0] : null
  };
}

function findByName(list, value) {
  return list.find((item) => normalizeSpreadsheetKey(item.title || item.name) === normalizeSpreadsheetKey(value));
}

function normalizeSheetAlias(value) {
  return normalizeSpreadsheetKey(value).replace(/[\s.]+/g, '');
}

function parseIssueConfigWorkbook(buffer) {
  const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true });
  const sheetName = workbook.SheetNames.find((name) => normalizeSheetAlias(name) === 'configissue');
  if (!sheetName) {
    const error = new Error("A planilha enviada nao possui a aba 'Config. issue'. Use a planilha modelo de configuracao de Issues.");
    error.status = 400;
    throw error;
  }
  const tableRows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: '', raw: false });
  const sections = [];
  let current = null;
  for (const row of tableRows) {
    const rowText = row.map((cell) => normalizeSpreadsheetKey(cell)).join(' ');
    const titleMatch =
      rowText.includes('catalogo de categorias') ||
      rowText.includes('catalogo de campos') ||
      rowText.includes('campos personalizados') ||
      rowText.includes('matriz de aplicacao') ||
      rowText.includes('listas de apoio');
    if (titleMatch) {
      current = { name: String(row.find((cell) => String(cell || '').trim()) || '').trim(), rows: [] };
      sections.push(current);
      continue;
    }
    if (current) current.rows.push(row);
  }
  return { sheetName, sections, tableRows };
}

function parseSectionTable(sectionRows, expectedHeaders = []) {
  const headerIdx = sectionRows.findIndex((row) => row.some((cell) => String(cell || '').trim()));
  if (headerIdx < 0) return [];
  const headers = sectionRows[headerIdx].map((h) => String(h || '').trim());
  const normalizedHeaders = headers.map((header) => normalizeSpreadsheetKey(header));
  if (expectedHeaders.length && !expectedHeaders.every((expected) => normalizedHeaders.some((header) => header.includes(expected)))) {
    return [];
  }
  return sectionRows
    .slice(headerIdx + 1)
    .map((row, idx) => {
      const obj = {};
      headers.forEach((h, colIdx) => {
        if (!h) return;
        obj[h] = normalizeCellValue(row[colIdx]);
      });
      obj.__line = headerIdx + idx + 2;
      return obj;
    })
    .filter((row) => {
      const lineValues = Object.entries(row).filter(([k]) => k !== '__line').map(([, v]) => normalizeSpreadsheetKey(v));
      if (!lineValues.some(Boolean)) return false;
      if (lineValues.some((value) => value.includes('codigo') && value.includes('categoria') && value.includes('tipo'))) return false;
      return true;
    });
}

function stripAccents(value) {
  return String(value || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

function normalizeComparableName(value) {
  return normalizeSpreadsheetKey(stripAccents(value));
}

function getRowValueByAliases(row, aliases) {
  const entries = Object.entries(row || {}).filter(([key]) => key !== '__line');
  for (const alias of aliases) {
    const normalizedAlias = normalizeSpreadsheetKey(alias).replace(/[._-]/g, ' ');
    const hit = entries.find(([key]) => {
      const normalizedKey = normalizeSpreadsheetKey(key).replace(/[._-]/g, ' ');
      return (
        normalizedKey === normalizedAlias ||
        normalizedKey.includes(normalizedAlias) ||
        normalizedAlias.includes(normalizedKey)
      );
    });
    if (hit) return normalizeCellValue(hit[1]);
  }
  return '';
}

function buildIssueConfigPreview(workbookData, context) {
  const getSection = (matcher) => workbookData.sections.find((s) => matcher(normalizeSpreadsheetKey(s.name || '')));
  const catalogSection = getSection((k) => k.includes('catalogo de categorias'));
  const fieldsSection = getSection((k) => k.includes('catalogo de campos') || k.includes('campos personalizados'));
  const matrixSection = getSection((k) => k.includes('matriz de aplicacao'));
  const categoriesAndTypes = parseSectionTable(catalogSection?.rows || [], ['categoria', 'tipo']);
  let customFields = parseSectionTable(fieldsSection?.rows || [], ['campo']);
  if (!customFields.length) {
    customFields = parseSectionTable(fieldsSection?.rows || [], []);
  }
  const matrixRows = parseSectionTable(matrixSection?.rows || [], ['campo']);
  const existingTypes = context.issueTypes.filter((x) => x.kind === 'type');
  const existingCategories = context.issueTypes.filter((x) => x.kind === 'category');
  const existingFields = context.fieldDefinitions || [];
  const categoryRows = [];
  const typeRows = [];
  const issues = [];
  for (const row of categoriesAndTypes) {
    const category = getRowValueByAliases(row, ['categoria acc', 'categoria']);
    const issueType = getRowValueByAliases(row, ['tipo de issue acc', 'tipo acc', 'tipo de issue', 'tipo']);
    const code = getRowValueByAliases(row, ['codigo', 'código', 'sigla']);
    if (category) {
      const exists = existingCategories.find((x) => normalizeComparableName(x.title) === normalizeComparableName(category));
      categoryRows.push({ line: row.__line, category, status: exists ? 'ja-existe' : 'pronto', action: exists ? 'Já existe' : 'Criar', observation: '' });
    }
    if (issueType || code) {
      const exists = existingTypes.find((x) =>
        normalizeComparableName(x.title) === normalizeComparableName(issueType) &&
        normalizeComparableName(x.category || '') === normalizeComparableName(category)
      );
      const observation = !category ? 'Categoria ACC obrigatória.' : !issueType ? 'Tipo ACC obrigatório.' : '';
      typeRows.push({ line: row.__line, code, category, issueType, when: row['Quando usar'] || '', status: observation ? 'erro' : (exists ? 'ja-existe' : 'pronto'), action: exists ? 'Já existe' : 'Criar', observation });
      if (observation) issues.push({ line: row.__line, section: 'Categorias/Tipos', status: 'erro', action: 'corrigir', reason: observation });
    }
  }
  const fieldRows = customFields.map((row) => {
    const name = getRowValueByAliases(row, ['campo', 'nome do campo']);
    const type = getRowValueByAliases(row, ['tipo acc sugerido', 'tipo acc', 'tipo']);
    const options = String(getRowValueByAliases(row, ['opções / lista', 'opcoes / lista', 'opções', 'opcoes']))
      .split(/[;,|]/).map((x) => x.trim()).filter(Boolean);
    const exists = existingFields.find((x) => normalizeComparableName(x.title || x.name) === normalizeComparableName(name));
    const listType = /lista/i.test(type);
    const observation = !name ? 'Campo obrigatório.' : !type ? 'Tipo ACC sugerido obrigatório.' : (listType && !options.length ? 'Campo de lista sem opções.' : '');
    return { line: row.__line, name, classification: getRowValueByAliases(row, ['classificação', 'classificacao']) || '', type, required: getRowValueByAliases(row, ['obrigatoriedade']) || '', applicability: getRowValueByAliases(row, ['aplicabilidade']) || '', options, status: observation ? 'erro' : (exists ? 'ja-existe' : 'pronto'), action: exists ? 'Já existe' : 'Criar', observation };
  });
  const matrix = matrixRows
    .map((row) => ({
      line: row.__line,
      code: getRowValueByAliases(row, ['código', 'codigo']) || '',
      issueType: getRowValueByAliases(row, ['tipo de issue', 'tipo de issue acc', 'tipo']) || '',
      category: getRowValueByAliases(row, ['categoria', 'categoria acc']) || '',
      field: getRowValueByAliases(row, ['campo']) || '',
      rule: (getRowValueByAliases(row, ['regra']) || '').toUpperCase(),
      interpretation: getRowValueByAliases(row, ['interpretação', 'interpretacao']) || ''
    }))
    .filter((row) => row.code || row.issueType || row.category || row.field);
  matrix.forEach((row) => {
    if (!['O', 'A', 'N', 'OBRIGATORIO', 'APLICAVEL', 'NAO'].includes(normalizeSpreadsheetKey(row.rule).replace(/\s+/g, ''))) {
      issues.push({ line: row.line, section: 'Matriz', status: 'alerta', action: 'revisar', reason: `Regra '${row.rule || '-'}' fora do padrão O/A/N.` });
    }
  });
  fieldRows.filter((row) => row.observation).forEach((row) => {
    issues.push({ line: row.line, section: 'Campos Personalizados', status: 'erro', action: 'corrigir', reason: row.observation });
  });
  return { sheetName: workbookData.sheetName, categories: categoryRows, types: typeRows, fields: fieldRows, matrix, issues, summary: { categories: new Set(categoryRows.map((r) => normalizeSpreadsheetKey(r.category))).size, types: typeRows.filter((r) => r.issueType).length, fields: fieldRows.filter((r) => r.name).length, matrix: matrix.length, errors: issues.filter((r) => r.status === 'erro').length, warnings: issues.filter((r) => r.status === 'alerta').length, ignored: 0 } };
}

function mapFieldTypeToAccType(typeName) {
  const key = normalizeSpreadsheetKey(typeName || '');
  if (!key) return 'text';
  if (key.includes('lista') || key.includes('select') || key.includes('dropdown')) return 'list';
  if (key.includes('numero') || key.includes('number')) return 'numeric';
  if (key.includes('data') || key.includes('date')) return 'date';
  if (key.includes('sim') && key.includes('nao')) return 'boolean';
  if (key.includes('multilinha') || key.includes('textarea')) return 'paragraph';
  return 'text';
}

app.post('/api/projects/:projectId/eap-issues/preview', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const workbookData =
      normalizeUploadedWorkbookData(req.body.workbookData) || parseEapWorkbook(decodeBase64Workbook(req.body.fileBase64));
    const context = await getIssueCreationContext(req, projectId);
    const rows = buildEapIssuePreviewRowsV2(workbookData, context);
    console.info('[EAP preview]', {
      projectId,
      sheetName: workbookData.sheetName,
      headerRow: workbookData.headerRowIndex,
      friendlyHeaderRow: workbookData.friendlyHeaderRowIndex,
      columns: (workbookData.columns || []).map((column) => column.issueField || column.friendlyField || column.label),
      customFieldsFound: context.fieldDefinitions.map((field) => field.name)
    });

    res.json({
      fileName: req.body.fileName || 'Planilha EAP',
      sheetName: workbookData.sheetName,
      headerRowIndex: workbookData.headerRowIndex,
      friendlyHeaderRowIndex: workbookData.friendlyHeaderRowIndex ?? -1,
      columns: (workbookData.columns || []).map((column) => column.issueField || column.friendlyField || column.label || column.key),
      workbookData,
      totalRows: rows.length,
      readyRows: rows.filter((row) => row.validation === 'pronto').length,
      warningRows: rows.filter((row) => row.validation === 'incompleto').length,
      duplicateRows: rows.filter((row) => row.validation === 'duplicado').length,
      errorRows: rows.filter((row) => row.validation === 'erro').length,
      rows,
      manualReference: {
        message: 'Validacao da aba ACS Build baseada nos campos obrigatorios do template.'
      }
    });
  } catch (error) {
    next(error);
  }
});

app.post('/api/projects/:projectId/eap-issues/preview-file', express.raw({ type: '*/*', limit: '25mb' }), async (req, res, next) => {
  try {
    const fileBuffer = Buffer.isBuffer(req.body) ? req.body : Buffer.from(req.body || []);

    if (!fileBuffer.length) {
      const error = new Error('Selecione uma planilha Excel para importar.');
      error.statusCode = 400;
      throw error;
    }

    const projectId = getAccProjectId(req.params.projectId);
    const workbookData = parseEapWorkbook(fileBuffer);
    const context = await getIssueCreationContext(req, projectId);
    const rows = buildEapIssuePreviewRowsV2(workbookData, { ...context, previewOnly: true });
    const encodedFileName = req.get('x-file-name') || '';
    const fileName = encodedFileName ? decodeURIComponent(encodedFileName) : 'Planilha EAP';

    res.json({
      fileName,
      sheetName: workbookData.sheetName,
      headerRowIndex: workbookData.headerRowIndex,
      friendlyHeaderRowIndex: workbookData.friendlyHeaderRowIndex ?? -1,
      columns: (workbookData.columns || []).map((column) => column.issueField || column.friendlyField || column.label || column.key),
      workbookData,
      fieldDefinitions: context.fieldDefinitions || [],
      users: context.users || [],
      issueTypes: context.issueTypes || [],
      totalRows: rows.length,
      readyRows: rows.filter((row) => row.validation === 'pronto').length,
      warningRows: rows.filter((row) => row.validation === 'incompleto').length,
      duplicateRows: rows.filter((row) => row.validation === 'duplicado').length,
      errorRows: rows.filter((row) => row.validation === 'erro').length,
      rows,
      manualReference: {
        message: 'Validacao da aba ACS Build baseada nos campos obrigatorios do template.'
      }
    });
  } catch (error) {
    next(error);
  }
});

app.post('/api/projects/:projectId/im-issues/preview-create-file', express.raw({ type: '*/*', limit: '25mb' }), async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const workbookData = parseIssueConfigWorkbook(Buffer.isBuffer(req.body) ? req.body : Buffer.from(req.body || []));
    const context = await getIssueCreationContext(req, projectId);
    const preview = buildIssueConfigPreview(workbookData, context);
    const hasCatalogIssue = preview.summary.categories === 0 || preview.summary.types === 0;
    res.json({ workbookData, ...preview, message: hasCatalogIssue ? 'Foram encontrados problemas na leitura de categorias/tipos. Prévia gerada. Revise antes de aplicar no ACC.' : "Seções identificadas com sucesso. Prévia gerada. Revise antes de aplicar no ACC." });
  } catch (error) { next(error); }
});

app.post('/api/projects/:projectId/im-issues/create', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const workbookData = req.body.workbookData;
    if (!workbookData) throw new Error('Prévia não encontrada. Recarregue a planilha.');
    console.info('[IM create] Iniciando aplicação da configuração de issues', {
      projectId,
      hasWorkbookData: Boolean(workbookData),
      sheetName: workbookData?.sheetName || ''
    });
    let context = await getIssueCreationContext(req, projectId);
    let me = null;
    try {
      me = await callApsApi(req, `/construction/admin/v1/projects/${projectId}/users/me`);
    } catch (error) {
      const usersResult = await callApsApi(req, `/construction/admin/v1/projects/${projectId}/users?limit=200`);
      me = (usersResult.results || []).find((user) =>
        String(user.autodeskId || user.id || '').trim() === String(req.session?.autodeskUser?.sub || '').trim()
      ) || null;
    }
    const permissions = Array.isArray(me?.permissions) ? me.permissions.map((x) => normalizeSpreadsheetKey(x)) : [];
    const roleName = normalizeSpreadsheetKey(me?.role || me?.projectRole || me?.roleName || '');
    const canCreateSettings =
      permissions.some((x) => x.includes('issues.new') || x.includes('issues.manage') || x.includes('issues.admin')) ||
      ['project admin', 'administrador', 'admin', 'account admin'].some((x) => roleName.includes(normalizeSpreadsheetKey(x)));
    if (!canCreateSettings) {
      const error = new Error('Seu usuário não possui permissão para criar configurações de Issues neste projeto ACC. Solicite acesso de administrador do projeto ou permissão issues.new.');
      error.status = 403;
      throw error;
    }
    const previewRows = buildIssueConfigPreview(workbookData, context);
    console.info('[IM create] Prévia consolidada', {
      categories: previewRows.categories.length,
      types: previewRows.types.length,
      fields: previewRows.fields.length,
      matrix: previewRows.matrix.length,
      errors: previewRows.summary?.errors || 0,
      warnings: previewRows.summary?.warnings || 0
    });
    const results = [];
    const applyAudit = [];
    const createdTypeIdsByName = new Map();
    const createdCategoryIdsByName = new Map();
    const createdCategoryTitles = new Set();
    const createdTypeTitles = new Set();
    const createdFieldTitles = new Set();
    for (const row of previewRows.categories) {
      if (row.action !== 'Criar') { results.push({ entity: 'categoria', line: row.line, name: row.category, status: 'ja-existe' }); continue; }
      try {
        const payload = { title: row.category };
        applyAudit.push({ entity: 'categoria', line: row.line, operation: 'create', endpoint: `/construction/issues/v1/projects/${projectId}/issue-types`, payload });
        const created = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-types`, { method: 'POST', body: payload });
        createdCategoryIdsByName.set(normalizeComparableName(row.category), created?.id || '');
        createdCategoryTitles.add(row.category);
        results.push({ entity: 'categoria', line: row.line, name: row.category, status: 'criado', accId: created?.id || '' });
      } catch (error) {
        results.push({ entity: 'categoria', line: row.line, name: row.category, status: 'erro', message: error.message });
      }
    }
    for (const row of previewRows.types) {
      if (row.action !== 'Criar') { results.push({ entity: 'tipo', line: row.line, name: row.issueType, status: 'ja-existe' }); continue; }
      if (row.observation) { results.push({ entity: 'tipo', line: row.line, name: row.issueType, status: 'erro', message: row.observation }); continue; }
      try {
        const refreshed = await getIssueCreationContext(req, projectId);
        context = refreshed;
        const parentCategory = context.issueTypes.find((x) => normalizeComparableName(x.title) === normalizeComparableName(row.category));
        const payload = { title: row.issueType, categoryId: parentCategory?.id || createdCategoryIdsByName.get(normalizeComparableName(row.category)) || undefined };
        applyAudit.push({ entity: 'tipo', line: row.line, operation: 'create', endpoint: `/construction/issues/v1/projects/${projectId}/issue-types`, payload });
        const created = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-types`, { method: 'POST', body: payload });
        createdTypeIdsByName.set(normalizeComparableName(row.issueType), created?.id || '');
        createdTypeTitles.add(row.issueType);
        results.push({ entity: 'tipo', line: row.line, name: row.issueType, status: 'criado', accId: created?.id || '' });
      } catch (error) {
        results.push({ entity: 'tipo', line: row.line, name: row.issueType, status: 'erro', message: error.message });
      }
    }
    for (const row of previewRows.fields) {
      if (row.action !== 'Criar') { results.push({ entity: 'campo', line: row.line, name: row.name, status: 'ja-existe' }); continue; }
      if (row.observation) { results.push({ entity: 'campo', line: row.line, name: row.name, status: 'erro', message: row.observation }); continue; }
      const payload = { title: row.name, description: row.classification || '', type: mapFieldTypeToAccType(row.type), allowedValues: row.options };
      try {
        applyAudit.push({ entity: 'campo', line: row.line, operation: 'create', endpoint: `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions`, payload });
        const created = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions`, { method: 'POST', body: payload });
        createdFieldTitles.add(row.name);
        const mappingResults = [];
        const mappedTypes = (previewRows.matrix || []).filter((matrixRow) => normalizeSpreadsheetKey(matrixRow.field) === normalizeSpreadsheetKey(row.name));
        context = await getIssueCreationContext(req, projectId);
        const latestTypes = context.issueTypes.filter((item) => item.kind === 'type' || item.kind === 'subtype');
        for (const matrixRow of mappedTypes) {
          const issueType = latestTypes.find((x) => normalizeComparableName(x.title) === normalizeComparableName(matrixRow.issueType));
          const mappedItemId = issueType?.id || createdTypeIdsByName.get(normalizeComparableName(matrixRow.issueType));
          if (!mappedItemId) {
            mappingResults.push(`Falha no mapeamento ${matrixRow.issueType}: tipo não encontrado no ACC após criação.`);
            continue;
          }
          try {
            const mappingPayload = { issueAttributeDefinitionId: created?.id, mappedItemId, itemType: issueType?.kind === 'subtype' ? 'issueSubtype' : 'issueType', required: ['o', 'obrigatorio', 'mandatory', 'required'].includes(normalizeSpreadsheetKey(matrixRow.rule || '')) };
            applyAudit.push({ entity: 'mapeamento', line: matrixRow.line, operation: 'create', endpoint: `/construction/issues/v1/projects/${projectId}/issue-attribute-mappings`, payload: mappingPayload });
            await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issue-attribute-mappings`, { method: 'POST', body: mappingPayload });
            mappingResults.push(`Mapeado em ${matrixRow.issueType}`);
          } catch (mappingError) {
            mappingResults.push(`Falha no mapeamento ${matrixRow.issueType}: ${mappingError.message}`);
          }
        }
        results.push({ entity: 'campo', line: row.line, name: row.name, status: 'criado', accId: created?.id || '', message: mappingResults.join(' | ') });
      } catch (error) {
        results.push({ entity: 'campo', line: row.line, name: row.name, status: 'erro', message: error.message });
      }
    }
    const createdRows = results.filter((row) => row.status === 'criado').length;
    const summaryByEntity = ['categoria', 'tipo', 'campo'].reduce((acc, entity) => {
      const rows = results.filter((row) => row.entity === entity);
      acc[entity] = {
        created: rows.filter((row) => row.status === 'criado').length,
        existing: rows.filter((row) => row.status === 'ja-existe').length,
        errors: rows.filter((row) => row.status === 'erro').length
      };
      return acc;
    }, {});
    const persistedContext = await getIssueCreationContext(req, projectId);
    const persistedChecks = {
      categories: Array.from(createdCategoryTitles).map((name) => ({
        name,
        persisted: persistedContext.issueTypes.some((item) => item.kind === 'category' && normalizeComparableName(item.title) === normalizeComparableName(name))
      })),
      types: Array.from(createdTypeTitles).map((name) => ({
        name,
        persisted: persistedContext.issueTypes.some((item) => (item.kind === 'type' || item.kind === 'subtype') && normalizeComparableName(item.title) === normalizeComparableName(name))
      })),
      fields: Array.from(createdFieldTitles).map((name) => ({
        name,
        persisted: persistedContext.fieldDefinitions.some((item) => normalizeComparableName(item.name || item.title) === normalizeComparableName(name))
      }))
    };
    const notPersisted = [...persistedChecks.categories, ...persistedChecks.types, ...persistedChecks.fields].filter((item) => !item.persisted);
    if (notPersisted.length) {
      console.warn('[IM create] Itens não confirmados no ACC após criação', {
        projectId,
        notPersisted
      });
    } else {
      console.info('[IM create] Persistência confirmada no ACC para todos os itens criados', {
        projectId,
        createdCategories: persistedChecks.categories.length,
        createdTypes: persistedChecks.types.length,
        createdFields: persistedChecks.fields.length
      });
    }
    console.info('[IM create] Auditoria da etapa apply', {
      projectId,
      operations: applyAudit.length,
      sample: applyAudit.slice(0, 10)
    });
    res.json({
      totalRows: results.length,
      createdRows,
      errorRows: results.filter((row) => row.status === 'erro').length,
      summaryByEntity,
      persistedChecks,
      persistenceVerified: notPersisted.length === 0,
      applyAudit,
      results,
      message: createdRows
        ? (notPersisted.length ? 'Configuração aplicada parcialmente: alguns itens não foram confirmados no ACC. Revise os resultados.' : 'Configuração de Issues aplicada no ACC com sucesso e persistência confirmada.')
        : 'Nenhum item novo foi criado. Verifique itens já existentes e erros.'
    });
  } catch (error) { next(error); }
});

app.post('/api/projects/:projectId/im-issues/preview-update-file', express.raw({ type: '*/*', limit: '25mb' }), async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const workbookData = parseEapWorkbookV2(Buffer.isBuffer(req.body) ? req.body : Buffer.from(req.body || []));
    const issuesResult = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues?limit=200`);
    const existing = (issuesResult.results || []).map(mapIssue);
    const rows = workbookData.rows.map((row, idx) => {
      const item = readImRowValues(row);
      const updateFlag = normalizeSpreadsheetKey(findColumnValue(row.values, ['atualizar?', 'atualizar', 'acao', 'ação']));
      const eligible = updateFlag !== 'nao' && updateFlag !== 'não';
      const { matches, found } = matchImIssue(existing, item);
      const errors = !eligible
        ? []
        : matches.length > 1
          ? ['Issue duplicado/ambíguo.']
          : found
            ? []
            : ['Issue não encontrado com correspondência segura.'];
      return {
        id: `${row.rowNumber}-${idx}`,
        line: row.rowNumber,
        title: found?.title || item.title,
        issueId: found?.id || '',
        validation: !eligible ? 'ignorado' : errors.length ? 'erro' : 'pronto para atualizar',
        errors,
        warnings: !eligible ? ['Ignorado por “Atualizar? = Não”.'] : [],
        current: found ? { status: found.status || '', dueDate: found.dueDate || '', assignee: found.assignedTo || '', priority: found.priority || '' } : null,
        proposed: { status: item.status || '', dueDate: item.dueDate || '', assignee: item.assignee || '', priority: item.priority || '' }
      };
    });
    res.json({ workbookData, rows, totalRows: rows.length, readyRows: rows.filter((r) => r.validation === 'pronto para atualizar').length });
  } catch (error) { next(error); }
});

app.post('/api/projects/:projectId/im-issues/update', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const workbookData = normalizeUploadedWorkbookData(req.body.workbookData);
    const allowClearEmpty = Boolean(req.body.allowClearEmpty);
    const issuesResult = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues?limit=200`);
    const existing = (issuesResult.results || []).map(mapIssue);
    const usersResult = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/users?limit=200`).catch(() => ({ results: [] }));
    const users = (usersResult.results || []).map(mapProjectUser);
    const definitionsResult = await callApsApi(
      req,
      `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions?limit=200`
    ).catch(() => ({ results: [] }));
    const definitions = Array.isArray(definitionsResult.results) ? definitionsResult.results : [];
    const definitionLookup = new Map(
      definitions.map((definition) => [normalizeSpreadsheetKey(definition.name || definition.title || definition.displayName || ''), definition])
    );
    const results = [];
    for (const row of workbookData.rows) {
      const item = readImRowValues(row);
      const updateFlag = normalizeSpreadsheetKey(findColumnValue(row.values, ['atualizar?', 'atualizar', 'acao', 'ação']));
      if (updateFlag === 'nao' || updateFlag === 'não') {
        results.push({ line: row.rowNumber, status: 'ignorado', message: 'Ignorado por “Atualizar? = Não”.' });
        continue;
      }
      const { matches, found } = matchImIssue(existing, item);
      if (matches.length > 1) {
        results.push({ line: row.rowNumber, status: 'erro', message: 'Issue duplicado/ambíguo.' });
        continue;
      }
      if (!found) { results.push({ line: row.rowNumber, status: 'erro', message: 'Sem correspondência segura.' }); continue; }
      const patch = {};
      if (item.title || allowClearEmpty) patch.title = item.title;
      if (item.description || allowClearEmpty) patch.description = item.description;
      if (item.status || allowClearEmpty) patch.status = item.status;
      if (item.dueDate || allowClearEmpty) patch.dueDate = item.dueDate;
      if (item.priority || allowClearEmpty) patch.priority = item.priority;
      if (item.assignee || allowClearEmpty) {
        const assignee = findProjectUser(users, item.assignee);
        if (assignee) {
          patch.assignedTo = assignee.id;
          patch.assignedToType = 'user';
        } else if (item.assignee) {
          results.push({ line: row.rowNumber, status: 'aviso', issueId: found.id, message: `Responsável "${item.assignee}" não encontrado no projeto.` });
        }
      }
      const customAttributes = Object.entries(item.customFields || {})
        .map(([name, value]) => {
          const definition = definitionLookup.get(normalizeSpreadsheetKey(name));
          if (!definition) return null;
          if (!value && !allowClearEmpty) return null;
          return {
            attributeDefinitionId: String(definition.id || definition.attributeDefinitionId || '').trim(),
            value: value || ''
          };
        })
        .filter((attribute) => attribute?.attributeDefinitionId);
      if (customAttributes.length) {
        patch.customAttributes = customAttributes;
      }
      if (Object.keys(patch).length === 0) {
        results.push({ line: row.rowNumber, status: 'ignorado', issueId: found.id, message: 'Linha sem campos elegíveis para atualização.' });
        continue;
      }
      await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues/${encodeURIComponent(found.id)}`, { method: 'PATCH', body: patch });
      results.push({ line: row.rowNumber, status: 'atualizado', issueId: found.id });
    }
    res.json({ results });
  } catch (error) { next(error); }
});

app.post('/api/projects/:projectId/eap-issues/create', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const dryRun = req.body.dryRun !== false;
    const submittedRows = Array.isArray(req.body.rows) ? req.body.rows : [];
    const submittedWorkbookData = normalizeUploadedWorkbookData(req.body.workbookData);
    const context = submittedWorkbookData ? await getIssueCreationContext(req, projectId) : null;
    const rows = submittedWorkbookData ? buildEapIssuePreviewRowsV2(submittedWorkbookData, context) : submittedRows;
    if (submittedWorkbookData && context) {
      console.info('[EAP create validation]', {
        projectId,
        totalRows: rows.length,
        ready: rows.filter((row) => row.validation === 'pronto').length,
        warnings: rows.filter((row) => row.validation === 'incompleto').length,
        errors: rows.filter((row) => row.validation === 'erro').length,
        customFieldsFound: context.fieldDefinitions.map((field) => field.name)
      });
    }
    const results = [];
    let issueTypesCache = null;

    async function resolveIssueTypeForPayload(payload) {
      if (payload.issueTypeId || payload.issueSubtypeId || !payload.issueTypeName) return payload;

      if (!issueTypesCache) {
        const issueTypesResult = await callApsApi(
          req,
          `/construction/issues/v1/projects/${projectId}/issue-types?include=subtypes&limit=200`
        );
        issueTypesCache = listIssueTypes(issueTypesResult);
      }

      const selectedIssueType = findIssueTypeByName(issueTypesCache, payload.issueTypeName);
      if (!selectedIssueType) {
        const error = new Error(`Tipo de Issue nao encontrado no ACC para "${payload.issueTypeName}".`);
        error.status = 400;
        throw error;
      }

      return {
        ...payload,
        issueTypeId: selectedIssueType.kind === 'type' ? selectedIssueType.id : payload.issueTypeId,
        issueSubtypeId: selectedIssueType.kind === 'subtype' ? selectedIssueType.id : payload.issueSubtypeId
      };
    }

    for (const row of rows) {
      if (row.validation === 'erro') {
        results.push({ line: row.line, eapCode: row.eapCode, title: row.title, status: 'erro', message: row.errors?.join(' ') || 'Linha com erro crítico.' });
        continue;
      }

      if (row.validation === 'duplicado') {
        results.push({ line: row.line, eapCode: row.eapCode, title: row.title, status: 'duplicado', message: 'Issue já existe no projeto.' });
        continue;
      }

      if (!row?.payload || !row.title) {
        results.push({
          line: row.line,
          eapCode: row.eapCode,
          title: row.title,
          status: 'revisar',
          message: 'Linha sem dados suficientes para criar Issue.'
        });
        continue;
      }

      if (dryRun) {
        results.push({ line: row.line, eapCode: row.eapCode, title: row.title, status: 'simulado', message: 'Modo teste: nenhum Issue foi criado.' });
        continue;
      }

      try {
        const payload = await resolveIssueTypeForPayload(row.payload);
        const { issueTypeName, ...payloadForAps } = payload;
        const cleanPayload = Object.fromEntries(
          Object.entries(payloadForAps).filter(([, value]) => value !== '' && value !== undefined && value !== null)
        );
        const issue = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues`, {
          method: 'POST',
          body: {
            ...cleanPayload,
            status: payload.status || 'open',
            published: true
          }
        });
        const mappedIssue = mapIssue(issue);
        let customUpdateMessage = '';
        if (Array.isArray(payload.customAttributes) && payload.customAttributes.length > 0 && mappedIssue.id) {
          await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues/${encodeURIComponent(mappedIssue.id)}`, {
            method: 'PATCH',
            body: {
              customAttributes: payload.customAttributes
            }
          })
            .then(() => {
              customUpdateMessage = ` ${payload.customAttributes.length} campos personalizados enviados.`;
            })
            .catch((updateError) => {
              customUpdateMessage = ` Issue criado, mas alguns campos personalizados precisam de revisao: ${formatApsError(updateError.message || '')}`;
            });
        }
        results.push({
          line: row.line,
          eapCode: row.eapCode,
          title: row.title,
          status: 'criado',
          issueId: mappedIssue.id,
          issueLink: issue.links?.webView?.href || issue.links?.self?.href || '',
          createdAt: new Date().toISOString(),
          nativeFieldsFilled: row.nativeFieldsFilled || [],
          customFieldsFilled: row.customFieldsFilled || [],
          customFieldsMissing: row.customFieldsMissing || [],
          message: `Issue criado no ACC.${customUpdateMessage}`
        });
      } catch (issueError) {
        results.push({
          line: row.line,
          eapCode: row.eapCode,
          title: row.title,
          status: 'erro',
          message: formatApsError(issueError.message || 'Falha ao criar Issue.')
        });
      }
    }

    res.json({ dryRun, results });
  } catch (error) {
    next(error);
  }
});

app.post('/api/projects/:projectId/issues', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);

    if (!req.body.title) {
      res.status(400).json({ message: 'Informe um titulo para criar a solicitacao.' });
      return;
    }

    const payload = {
      title: req.body.title,
      description: req.body.description || '',
      status: 'open',
      published: true
    };

    if (req.body.dueDate) {
      payload.dueDate = req.body.dueDate;
    }

    if (req.body.assignedTo) {
      payload.assignedTo = req.body.assignedTo;
      payload.assignedToType = 'user';
    }

    if (req.body.issueSubtypeId) {
      payload.issueSubtypeId = req.body.issueSubtypeId;
    } else if (req.body.issueTypeId) {
      payload.issueTypeId = req.body.issueTypeId;
    } else {
      const issueTypesResult = await callApsApi(
        req,
        `/construction/issues/v1/projects/${projectId}/issue-types?include=subtypes&limit=200`
      );
      const issueTypes = listIssueTypes(issueTypesResult);
      const selectedIssueType = findIssueTypeByName(issueTypes, req.body.issueTypeName);

      if (!selectedIssueType) {
        res.status(400).json({
          message: `Nao encontrei no projeto um tipo/categoria de issue com o nome "${req.body.issueTypeName}". Confira a configuracao de Issues no ACC.`
        });
        return;
      }

      if (selectedIssueType.kind === 'subtype') {
        payload.issueSubtypeId = selectedIssueType.id;
      } else {
        payload.issueTypeId = selectedIssueType.id;
      }
    }

    if (Array.isArray(req.body.customAttributes) && req.body.customAttributes.length > 0) {
      const definitionsResult = await callApsApi(
        req,
        `/construction/issues/v1/projects/${projectId}/issue-attribute-definitions?limit=200`
      ).catch(() => ({ results: [] }));
      const definitions = Array.isArray(definitionsResult.results)
        ? definitionsResult.results
        : Array.isArray(definitionsResult.data)
          ? definitionsResult.data
          : [];
      const readOnlyIds = new Set(
        definitions
          .filter((definition) => definition?.readOnly === true)
          .map((definition) => String(definition.id || definition.attributeDefinitionId || '').trim())
          .filter(Boolean)
      );

      payload.customAttributes = req.body.customAttributes
        .map((attribute) => ({
          attributeDefinitionId: String(attribute?.attributeDefinitionId || '').trim(),
          value: attribute?.value ?? ''
        }))
        .filter((attribute) => attribute.attributeDefinitionId && !readOnlyIds.has(attribute.attributeDefinitionId));

      if (payload.customAttributes.length === 0) {
        delete payload.customAttributes;
      }
    }

    const issue = await callApsApi(req, `/construction/issues/v1/projects/${projectId}/issues`, {
      method: 'POST',
      body: payload
    });

    res.status(201).json(mapIssue(issue));
  } catch (error) {
    next(error);
  }
});

app.patch('/api/projects/:projectId/issues/:issueId', async (req, res, next) => {
  try {
    const projectId = getAccProjectId(req.params.projectId);
    const allowedFields = ['title', 'description', 'status', 'dueDate', 'priority', 'assignedTo', 'assignedToType', 'watchers', 'followers', 'issueTypeId', 'issueSubtypeId', 'customAttributes'];
    const payload = {};

    for (const field of allowedFields) {
      if (req.body[field] !== undefined) {
        if (field === 'followers' && req.body.watchers === undefined) {
          payload.watchers = Array.isArray(req.body[field]) ? req.body[field].filter(Boolean) : req.body[field] || null;
        } else if (field !== 'followers') {
          payload[field] = Array.isArray(req.body[field]) ? req.body[field].filter(Boolean) : req.body[field] || null;
        }
      }
    }

    if (Object.keys(payload).length === 0) {
      res.status(400).json({ message: 'Nenhum campo permitido foi enviado para atualizar.' });
      return;
    }

    // Atualizacao controlada: por enquanto o app permite alterar titulo, status e prazo.
    const issue = await callApsApi(
      req,
      `/construction/issues/v1/projects/${projectId}/issues/${encodeURIComponent(req.params.issueId)}`,
      {
        method: 'PATCH',
        body: payload
      }
    );

    res.json(mapIssue(issue));
  } catch (error) {
    next(error);
  }
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok', app: 'g5-instrumentos-controle' });
});

app.use(express.static(distPath));

app.get(/^(?!\/api).*/, (req, res) => {
  res.sendFile(path.join(distPath, 'index.html'));
});

app.use((error, req, res, next) => {
  const status = error.status || 500;

  if (status >= 500) {
    console.error(error);
  } else if (status === 401) {
    console.warn(`[auth] ${error.message || 'Sessao Autodesk nao autenticada.'}`);
  } else {
    console.warn(error.message || error);
  }

  res.status(status).json({
    message: formatApsError(error.message || 'Algo deu errado. Tente novamente.')
  });
});

app.listen(port, () => {
  console.log(`Backend rodando em http://localhost:${port}`);
  console.log(`Frontend esperado em ${frontendUrl}`);
});
