export interface AdoCredentials {
  organization: string
  pat: string
}

export interface AdoWorkItem {
  id: number
  title: string
  workItemType: string
  state: string
  project: string
  organization: string
  assignedTo: string
  createdBy: string
  createdDate?: string
  changedBy: string
  changedDate?: string
  description: string
  webUrl: string
}

export interface AdoComment {
  id: number
  text: string
  createdBy: string
  createdDate: string
  modifiedDate?: string
}

interface WiqlResponse {
  workItems?: Array<{ id: number; url: string }>
}

interface WorkItemsBatchResponse {
  count: number
  value: Array<{
    id: number
    fields: Record<string, unknown>
    _links?: { html?: { href?: string } }
  }>
}

function authHeader(pat: string): string {
  return `Basic ${Buffer.from(`:${pat}`).toString('base64')}`
}

/** Accepts org name or full Azure DevOps URL and returns the org slug. */
export function normalizeOrganization(input: string): string {
  const raw = input.trim()
  if (!raw) {
    throw new Error('Organization is required.')
  }

  const withProtocol = /^https?:\/\//i.test(raw) ? raw : `https://${raw}`

  try {
    const url = new URL(withProtocol)
    const host = url.hostname.toLowerCase()

    if (host.endsWith('.visualstudio.com')) {
      const org = host.split('.')[0]
      if (org && org !== 'www') return validateOrgSlug(org)
    }

    if (host === 'dev.azure.com') {
      const segment = url.pathname.split('/').filter(Boolean)[0]
      if (segment) return validateOrgSlug(decodeURIComponent(segment))
    }
  } catch (e) {
    if (e instanceof Error && e.message.includes('Organization')) throw e
    // Fall through
  }

  let value = raw
    .replace(/^https?:\/\//i, '')
    .replace(/^dev\.azure\.com\//i, '')
    .replace(/\.visualstudio\.com.*$/i, '')

  value = value.split(/[/?#]/)[0]?.trim() ?? ''

  if (!value) {
    throw new Error(
      'Could not parse organization. Paste your ADO URL, e.g. https://dev.azure.com/my-org'
    )
  }

  return validateOrgSlug(value)
}

function validateOrgSlug(org: string): string {
  // Accidentally typed spaces → strip them (e.g. "Sekisui Diagnostics" → "SekisuiDiagnostics")
  const slug = org.trim().replace(/\s+/g, '')
  if (!slug) {
    throw new Error('Organization is required.')
  }
  if (!/^[A-Za-z0-9]([A-Za-z0-9-]*[A-Za-z0-9])?$/.test(slug)) {
    throw new Error(
      `Invalid organization "${slug}". Use only letters, numbers, and hyphens — copy it from your browser URL after https://dev.azure.com/`
    )
  }
  return slug
}

function orgBase(organization: string): string {
  const org = normalizeOrganization(organization)
  return `https://dev.azure.com/${encodeURIComponent(org)}`
}

function displayName(value: unknown): string {
  if (!value) return ''
  if (typeof value === 'string') return value
  if (typeof value === 'object' && value !== null && 'displayName' in value) {
    return String((value as { displayName?: string }).displayName ?? '')
  }
  return String(value)
}

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ')
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/\s+/g, ' ')
    .trim()
}

/** Keep markup for rich display; normalize empty content. */
function normalizeHtml(html: string): string {
  const trimmed = (html ?? '').trim()
  if (!trimmed || trimmed === '<div></div>' || trimmed === '<p></p>') return ''
  return trimmed
}

function toWorkItem(
  org: string,
  item: {
    id: number
    fields: Record<string, unknown>
    _links?: { html?: { href?: string } }
  }
): AdoWorkItem {
  const project = String(item.fields['System.TeamProject'] ?? '')
  const title = String(item.fields['System.Title'] ?? '')
  const descriptionRaw = String(
    item.fields['System.Description'] ?? item.fields['Microsoft.VSTS.TCM.ReproSteps'] ?? ''
  )

  return {
    id: item.id,
    title,
    workItemType: String(item.fields['System.WorkItemType'] ?? ''),
    state: String(item.fields['System.State'] ?? ''),
    project,
    organization: org,
    assignedTo: displayName(item.fields['System.AssignedTo']),
    createdBy: displayName(item.fields['System.CreatedBy']),
    createdDate: item.fields['System.CreatedDate']
      ? String(item.fields['System.CreatedDate'])
      : undefined,
    changedBy: displayName(item.fields['System.ChangedBy']),
    changedDate: item.fields['System.ChangedDate']
      ? String(item.fields['System.ChangedDate'])
      : undefined,
    description: normalizeHtml(descriptionRaw),
    webUrl:
      item._links?.html?.href ??
      `https://dev.azure.com/${org}/${encodeURIComponent(project)}/_workitems/edit/${item.id}`
  }
}

function friendlyHttpError(status: number, body: string, url: string): string {
  const snippet = body.replace(/\s+/g, ' ').slice(0, 240)
  if (status === 401 || status === 403) {
    return `Azure DevOps ${status}: Auth failed. Check your PAT (needs Work Items Read). Tried: ${url}`
  }
  if (status === 404) {
    return `Azure DevOps 404: Organization not found. Check the org name in Settings. Tried: ${url}`
  }
  return `Azure DevOps ${status}: ${snippet || 'Request failed'} (Tried: ${url})`
}

async function adoFetch<T>(
  credentials: AdoCredentials,
  path: string,
  init?: RequestInit
): Promise<T> {
  const base = orgBase(credentials.organization)
  const url = `${base}${path}`
  const res = await fetch(url, {
    ...init,
    headers: {
      Authorization: authHeader(credentials.pat),
      Accept: 'application/json',
      'Content-Type': 'application/json',
      ...(init?.headers ?? {})
    }
  })

  if (!res.ok) {
    const body = await res.text().catch(() => '')
    throw new Error(friendlyHttpError(res.status, body, url))
  }

  if (res.status === 204) return undefined as T
  return (await res.json()) as T
}

const EXCLUDED_STATES = [
  'Backlog',
  'Resolved',
  'Fixed',
  'Closed',
  'Done',
  'Removed',
  'Completed',
  'Cut',
  'Inactive',
  'Rejected'
] as const

export function isExcludedWorkItemState(state: string): boolean {
  const s = (state ?? '').trim().toLowerCase()
  if (!s) return false
  if (s.includes('backlog')) return true
  if (s === 'resolved' || s.startsWith('resolved ')) return true
  if (s === 'fixed' || s.startsWith('fixed ')) return true
  return EXCLUDED_STATES.some((name) => name.toLowerCase() === s)
}

async function fetchWorkItemsByIds(
  credentials: AdoCredentials,
  org: string,
  ids: number[]
): Promise<AdoWorkItem[]> {
  if (ids.length === 0) return []

  const chunks: number[][] = []
  for (let i = 0; i < ids.length; i += 200) {
    chunks.push(ids.slice(i, i + 200))
  }

  const results: AdoWorkItem[] = []
  for (const chunk of chunks) {
    const data = await adoFetch<WorkItemsBatchResponse>(
      credentials,
      `/_apis/wit/workitems?ids=${chunk.join(',')}&$expand=All&api-version=7.1`
    )
    for (const item of data.value ?? []) {
      results.push(toWorkItem(org, item))
    }
  }
  return results
}

export interface AdoIdentity {
  id: string
  displayName: string
  uniqueName: string
}

/** Who the PAT authenticates as — this is what WIQL @Me resolves to. */
export async function fetchAuthenticatedUser(
  credentials: AdoCredentials
): Promise<AdoIdentity> {
  const data = await adoFetch<{
    authenticatedUser?: {
      id?: string
      providerDisplayName?: string
      customDisplayName?: string
      properties?: Record<string, { $value?: string } | undefined>
    }
  }>(credentials, '/_apis/connectionData')

  const user = data.authenticatedUser
  const displayName =
    user?.customDisplayName?.trim() ||
    user?.providerDisplayName?.trim() ||
    'Unknown user'
  const uniqueName =
    user?.properties?.Account?.$value?.trim() ||
    user?.properties?.Mail?.$value?.trim() ||
    ''

  return {
    id: user?.id ?? '',
    displayName,
    uniqueName
  }
}

async function fetchWorkItemsAssignedToMe(
  credentials: AdoCredentials,
  org: string,
  ids: number[]
): Promise<AdoWorkItem[]> {
  if (ids.length === 0) return []

  const chunks: number[][] = []
  for (let i = 0; i < ids.length; i += 200) {
    chunks.push(ids.slice(i, i + 200))
  }

  // Trust WIQL `@Me` for assignment; drop excluded states after expand.
  const results: AdoWorkItem[] = []
  for (const chunk of chunks) {
    const data = await adoFetch<WorkItemsBatchResponse>(
      credentials,
      `/_apis/wit/workitems?ids=${chunk.join(',')}&$expand=All&api-version=7.1`
    )
    for (const item of data.value ?? []) {
      const state = String(item.fields['System.State'] ?? '')
      if (isExcludedWorkItemState(state)) continue
      results.push(toWorkItem(org, item))
    }
  }
  return results
}

/**
 * Active work items assigned to the PAT user (@Me in Azure DevOps).
 * Matches the Boards “Assigned to me” style query.
 */
export async function fetchMyWorkItems(
  credentials: AdoCredentials,
  projectNames?: string[]
): Promise<{ items: AdoWorkItem[]; asUser: AdoIdentity }> {
  const org = normalizeOrganization(credentials.organization)
  const asUser = await fetchAuthenticatedUser(credentials)

  // Same macro DevOps uses in the query editor: @Me
  const parts = [
    'SELECT [System.Id]',
    'FROM WorkItems',
    'WHERE [System.AssignedTo] = @Me',
    `AND [System.State] NOT IN (${EXCLUDED_STATES.map((s) => `'${s}'`).join(', ')})`
  ]

  const projects = (projectNames ?? []).map((p) => p.trim()).filter(Boolean)
  if (projects.length > 0) {
    const clauses = projects
      .map((p) => `[System.TeamProject] = '${p.replace(/'/g, "''")}'`)
      .join(' OR ')
    parts.push(`AND (${clauses})`)
  }

  parts.push('ORDER BY [System.ChangedDate] DESC')

  const query = await adoFetch<WiqlResponse>(
    credentials,
    '/_apis/wit/wiql?$top=2000&api-version=7.1',
    {
      method: 'POST',
      body: JSON.stringify({ query: parts.join(' ') })
    }
  )

  const ids = (query.workItems ?? []).map((w) => w.id)
  let items = await fetchWorkItemsAssignedToMe(credentials, org, ids)

  // If WIQL @Me returned nothing, retry with explicit uniqueName (some orgs/PAT edge cases)
  if (items.length === 0 && asUser.uniqueName) {
    const fallbackParts = [
      'SELECT [System.Id]',
      'FROM WorkItems',
      `WHERE [System.AssignedTo] = '${asUser.uniqueName.replace(/'/g, "''")}'`,
      `AND [System.State] NOT IN (${EXCLUDED_STATES.map((s) => `'${s}'`).join(', ')})`
    ]
    if (projects.length > 0) {
      const clauses = projects
        .map((p) => `[System.TeamProject] = '${p.replace(/'/g, "''")}'`)
        .join(' OR ')
      fallbackParts.push(`AND (${clauses})`)
    }
    fallbackParts.push('ORDER BY [System.ChangedDate] DESC')

    const fallback = await adoFetch<WiqlResponse>(
      credentials,
      '/_apis/wit/wiql?$top=2000&api-version=7.1',
      {
        method: 'POST',
        body: JSON.stringify({ query: fallbackParts.join(' ') })
      }
    )
    items = await fetchWorkItemsAssignedToMe(
      credentials,
      org,
      (fallback.workItems ?? []).map((w) => w.id)
    )
  }

  return { items, asUser }
}

export interface AdoProject {
  id: string
  name: string
  description: string
  state: string
}

export async function fetchProjects(
  credentials: AdoCredentials
): Promise<AdoProject[]> {
  const data = await adoFetch<{
    value?: Array<{
      id: string
      name: string
      description?: string
      state: string
    }>
  }>(credentials, '/_apis/projects?api-version=7.1&$top=1000&stateFilter=WellFormed')

  return (data.value ?? [])
    .map((p) => ({
      id: p.id,
      name: p.name,
      description: p.description ?? '',
      state: p.state
    }))
    .sort((a, b) => a.name.localeCompare(b.name))
}

export async function fetchWorkItem(
  credentials: AdoCredentials,
  id: number
): Promise<AdoWorkItem> {
  const org = normalizeOrganization(credentials.organization)
  const items = await fetchWorkItemsByIds(credentials, org, [id])
  if (!items[0]) throw new Error(`Work item ${id} not found.`)
  return items[0]
}

export async function updateWorkItem(
  credentials: AdoCredentials,
  id: number,
  patch: { title?: string; state?: string; description?: string }
): Promise<AdoWorkItem> {
  const ops: Array<{ op: string; path: string; value: string }> = []
  if (patch.title !== undefined) {
    ops.push({ op: 'add', path: '/fields/System.Title', value: patch.title })
  }
  if (patch.state !== undefined) {
    ops.push({ op: 'add', path: '/fields/System.State', value: patch.state })
  }
  if (patch.description !== undefined) {
    ops.push({
      op: 'add',
      path: '/fields/System.Description',
      value: patch.description
    })
  }

  if (ops.length === 0) return fetchWorkItem(credentials, id)

  await adoFetch(credentials, `/_apis/wit/workitems/${id}?api-version=7.1`, {
    method: 'PATCH',
    headers: {
      'Content-Type': 'application/json-patch+json'
    },
    body: JSON.stringify(ops)
  })

  return fetchWorkItem(credentials, id)
}

/** Discussion / comments for a work item (newest last). */
export async function fetchWorkItemComments(
  credentials: AdoCredentials,
  workItemId: number,
  project: string
): Promise<AdoComment[]> {
  const org = normalizeOrganization(credentials.organization)
  const projectSegment = encodeURIComponent(project)
  const data = await adoFetch<{
    totalCount?: number
    comments?: Array<{
      id: number
      text?: string
      createdDate?: string
      modifiedDate?: string
      createdBy?: unknown
    }>
  }>(
    credentials,
    `/${projectSegment}/_apis/wit/workItems/${workItemId}/comments?$top=200&order=asc&api-version=7.1-preview.4`
  )

  // org used only for clarity in errors via adoFetch base URL
  void org

  return (data.comments ?? []).map((c) => ({
    id: c.id,
    text: normalizeHtml(c.text ?? ''),
    createdBy: displayName(c.createdBy),
    createdDate: c.createdDate ?? '',
    modifiedDate: c.modifiedDate
  }))
}

/** Plain-text preview helper for lists / search. */
export function htmlToPlainText(html: string): string {
  return stripHtml(html)
}
