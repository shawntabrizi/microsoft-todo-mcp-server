import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
import { z } from "zod"
import { readFileSync, writeFileSync, existsSync } from "fs"
import { join } from "path"
import dotenv from "dotenv"
import { tokenManager } from "./token-manager.js"

// Load environment variables
dotenv.config()

// Log the current working directory
console.error("Current working directory:", process.cwd())

// Microsoft Graph API endpoints
const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const USER_AGENT = "microsoft-todo-mcp-server/1.0"
const LOG_PREVIEW_LIMIT = 4000

// Create server instance
const server = new McpServer({
  name: "mstodo",
  version: "1.0.0",
})

type GraphRequestErrorInfo = {
  url: string
  method: string
  status?: number
  responseBody?: string
  requestBody?: string
  requestId?: string
  clientRequestId?: string
  message: string
}

let lastGraphRequestError: GraphRequestErrorInfo | null = null

// Deduplication cache for non-idempotent requests (POST/PATCH/DELETE)
// Prevents duplicate task creation when MCP tools are double-invoked
const inFlightRequests = new Map<string, Promise<any>>()

function getRequestKey(method: string, url: string, body?: string): string {
  return `${method}:${url}:${body ?? ""}`
}

type RecurringDatePatchResult = {
  task?: Task
  error?: string
}

// Helper function for making Microsoft Graph API requests
async function makeGraphRequest<T>(url: string, token: string, method = "GET", body?: any): Promise<T | null> {
  const headers = {
    "User-Agent": USER_AGENT,
    Accept: "application/json",
    Authorization: `Bearer ${token}`,
    "Content-Type": "application/json",
  }
  const serializedBody = body && (method === "POST" || method === "PATCH") ? JSON.stringify(body) : undefined

  // Deduplicate non-idempotent requests by sharing in-flight promises
  if (method === "POST" || method === "PATCH" || method === "DELETE") {
    const cacheKey = getRequestKey(method, url, serializedBody)
    const existing = inFlightRequests.get(cacheKey)
    if (existing) {
      console.error(`Deduplicating ${method} request to ${url}`)
      return existing as Promise<T>
    }
    const promise = makeGraphRequestInner<T>(url, token, method, body, headers, serializedBody)
    inFlightRequests.set(cacheKey, promise)
    promise.finally(() => inFlightRequests.delete(cacheKey))
    return promise
  }

  return makeGraphRequestInner<T>(url, token, method, body, headers, serializedBody)
}

async function makeGraphRequestInner<T>(
  url: string,
  token: string,
  method: string,
  body: any,
  headers: Record<string, string>,
  serializedBody: string | undefined,
): Promise<T | null> {
  try {
    lastGraphRequestError = null

    const options: RequestInit = {
      method,
      headers,
    }

    if (serializedBody) {
      options.body = serializedBody
    }

    console.error(`${method} ${url}`)

    let response = await fetch(url, options)

    // If we get a 401, force a refresh and retry once.
    if (response.status === 401) {
      console.error("Got 401, attempting token refresh...")
      const newToken = await getAccessToken(true)
      if (newToken && newToken !== token) {
        headers.Authorization = `Bearer ${newToken}`
        response = await fetch(url, { ...options, headers })
      }
    }

    if (!response.ok) {
      const errorText = await response.text()
      lastGraphRequestError = extractGraphErrorInfo(url, method, serializedBody, response.status, errorText)

      console.error(`HTTP error! status: ${response.status}`)

      // Check for the specific MailboxNotEnabledForRESTAPI error
      if (errorText.includes("MailboxNotEnabledForRESTAPI")) {
        console.error(`
=================================================================
ERROR: MailboxNotEnabledForRESTAPI

The Microsoft To Do API is not available for personal Microsoft accounts
(outlook.com, hotmail.com, live.com, etc.) through the Graph API.

This is a limitation of the Microsoft Graph API, not an authentication issue.
Microsoft only allows To Do API access for Microsoft 365 business accounts.

You can still use Microsoft To Do through the web interface or mobile apps,
but API access is restricted for personal accounts.
=================================================================
        `)

        throw new Error(
          "Microsoft To Do API is not available for personal Microsoft accounts. See console for details.",
        )
      }

      throw new Error(`HTTP error! status: ${response.status}, body: ${errorText}`)
    }

    if (response.status === 204) {
      console.error("Response received: 204 No Content")
      return null
    }

    const responseText = await response.text()
    if (!responseText.trim()) {
      console.error("Response received: empty body")
      return null
    }

    const data = JSON.parse(responseText)
    return data as T
  } catch (error) {
    console.error("Error making Graph API request:", error)
    if (!lastGraphRequestError && error instanceof Error) {
      lastGraphRequestError = {
        url,
        method,
        requestBody: serializedBody,
        message: error.message,
      }
    }
    // Re-throw for DELETE so callers can distinguish failure from 204 success (both return null)
    if (method === "DELETE") {
      throw error
    }
    return null
  }
}

// Authentication helper using delegated flow with token manager
async function getAccessToken(forceRefresh = false): Promise<string | null> {
  try {
    console.error("getAccessToken called")

    const tokens = await tokenManager.getTokens({ forceRefresh })

    if (tokens) {
      console.error(`Successfully retrieved valid token`)
      return tokens.accessToken
    }

    console.error("No valid tokens available")
    return null
  } catch (error) {
    console.error("Error getting access token:", error)
    return null
  }
}

// Server configuration type
interface ServerConfig {
  accessToken?: string
  refreshToken?: string
  tokenFilePath?: string
}

// Function to check if the account is a personal Microsoft account
async function isPersonalMicrosoftAccount(): Promise<boolean> {
  try {
    const token = await getAccessToken()
    if (!token) return false

    const url = `${MS_GRAPH_BASE}/me`
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/json",
      },
    })

    if (!response.ok) {
      console.error(`Error getting user info: ${response.status}`)
      return false
    }

    const userData = await response.json()
    const email = userData.mail || userData.userPrincipalName || ""

    const personalDomains = ["outlook.com", "hotmail.com", "live.com", "msn.com", "passport.com"]
    const domain = email.split("@")[1]?.toLowerCase()

    if (domain && personalDomains.some((d) => domain.includes(d))) {
      console.error(`
=================================================================
WARNING: Personal Microsoft Account Detected

Your Microsoft account (${email}) appears to be a personal account.
Microsoft To Do API access is typically not available for personal accounts
through the Microsoft Graph API, only for Microsoft 365 business accounts.

You may encounter the "MailboxNotEnabledForRESTAPI" error when trying to
access To Do lists or tasks. This is a limitation of the Microsoft Graph API,
not an issue with your authentication or this application.
=================================================================
      `)
      return true
    }

    return false
  } catch (error) {
    console.error("Error checking account type:", error)
    return false
  }
}

// === Type definitions ===

interface TaskList {
  id: string
  displayName: string
  isOwner?: boolean
  isShared?: boolean
  wellknownListName?: string
}

interface DateTimeTimeZone {
  dateTime: string
  timeZone: string
}

interface RecurrencePattern {
  type: string
  interval: number
  month?: number
  dayOfMonth?: number
  daysOfWeek?: string[]
  firstDayOfWeek?: string
  index?: string
}

interface RecurrenceRange {
  type: string
  startDate: string
  endDate?: string
  recurrenceTimeZone?: string
  numberOfOccurrences?: number
}

interface PatternedRecurrence {
  pattern: RecurrencePattern
  range: RecurrenceRange
}

interface LinkedResource {
  id?: string
  webUrl?: string
  applicationName?: string
  displayName?: string
  externalId?: string
}

interface TaskFileAttachment {
  id: string
  name: string
  contentType?: string
  size?: number
  lastModifiedDateTime?: string
  contentBytes?: string
}

interface DeltaResponse<T> {
  value: T[]
  "@odata.count"?: number
  "@odata.nextLink"?: string
  "@odata.deltaLink"?: string
}

interface Task {
  id: string
  title: string
  status: string
  importance: string
  dueDateTime?: DateTimeTimeZone
  startDateTime?: DateTimeTimeZone
  completedDateTime?: DateTimeTimeZone
  reminderDateTime?: DateTimeTimeZone
  isReminderOn?: boolean
  recurrence?: PatternedRecurrence | null
  hasAttachments?: boolean
  createdDateTime?: string
  lastModifiedDateTime?: string
  bodyLastModifiedDateTime?: string
  body?: {
    content: string
    contentType: string
  }
  categories?: string[]
  linkedResources?: LinkedResource[]
  checklistItems?: ChecklistItem[]
}

interface ChecklistItem {
  id: string
  displayName: string
  isChecked: boolean
  checkedDateTime?: string
  createdDateTime?: string
}

interface UploadSession {
  uploadUrl: string
  expirationDateTime: string
  nextExpectedRanges: string[]
}

// === Zod schemas for tool parameters ===

const recurrenceSchema = z.object({
  pattern: z.object({
    type: z.string().describe("Recurrence type such as daily, weekly, absoluteMonthly, relativeMonthly"),
    interval: z.number().int().min(1).describe("Repeat interval"),
    month: z.number().int().min(1).max(12).optional(),
    dayOfMonth: z.number().int().min(1).max(31).optional(),
    daysOfWeek: z.array(z.string()).optional(),
    firstDayOfWeek: z.string().optional(),
    index: z.string().optional(),
  }),
  range: z.object({
    type: z.string().describe("Range type such as noEnd, endDate, or numbered"),
    startDate: z.string().describe("Start date in YYYY-MM-DD format"),
    endDate: z.string().optional().describe("End date in YYYY-MM-DD format"),
    recurrenceTimeZone: z.string().optional(),
    numberOfOccurrences: z.number().int().min(1).optional(),
  }),
})

type RecurrenceInput = z.infer<typeof recurrenceSchema>

const linkedResourceSchema = z.object({
  webUrl: z.string().optional().describe("Deep link to the linked item"),
  applicationName: z.string().optional().describe("Source application name"),
  displayName: z.string().optional().describe("Display title for the linked item"),
  externalId: z.string().optional().describe("External identifier from the source system"),
})

// === Helper functions ===

function buildDateTimeTimeZone(dateTime: string): DateTimeTimeZone {
  return {
    dateTime,
    timeZone: "UTC",
  }
}

function buildRecurrencePayload(recurrence: RecurrenceInput): PatternedRecurrence {
  const pattern: PatternedRecurrence["pattern"] = {
    type: recurrence.pattern.type,
    interval: recurrence.pattern.interval,
  }

  if (recurrence.pattern.month !== undefined) pattern.month = recurrence.pattern.month
  if (recurrence.pattern.dayOfMonth !== undefined) pattern.dayOfMonth = recurrence.pattern.dayOfMonth
  if (recurrence.pattern.daysOfWeek !== undefined) pattern.daysOfWeek = recurrence.pattern.daysOfWeek
  if (recurrence.pattern.firstDayOfWeek !== undefined) pattern.firstDayOfWeek = recurrence.pattern.firstDayOfWeek
  if (recurrence.pattern.index !== undefined) pattern.index = recurrence.pattern.index

  const range: PatternedRecurrence["range"] = {
    type: recurrence.range.type,
    startDate: recurrence.range.startDate,
  }

  if (recurrence.range.endDate !== undefined) range.endDate = recurrence.range.endDate
  if (recurrence.range.numberOfOccurrences !== undefined)
    range.numberOfOccurrences = recurrence.range.numberOfOccurrences
  if (recurrence.range.recurrenceTimeZone !== undefined) range.recurrenceTimeZone = recurrence.range.recurrenceTimeZone

  return { pattern, range }
}

function buildRecurrencePatchPayload(recurrence: RecurrenceInput): PatternedRecurrence {
  return {
    pattern: buildRecurrencePayload(recurrence).pattern,
    range: {} as PatternedRecurrence["range"],
  }
}

function buildRecurrencePatchPayloadFromExisting(recurrence: PatternedRecurrence): PatternedRecurrence {
  return {
    pattern: recurrence.pattern,
    range: {} as PatternedRecurrence["range"],
  }
}

function formatBodyForLog(body: unknown): string {
  if (body === undefined) return ""

  try {
    const serialized = JSON.stringify(body)
    return serialized.length > LOG_PREVIEW_LIMIT
      ? `${serialized.slice(0, LOG_PREVIEW_LIMIT)}... [truncated ${serialized.length - LOG_PREVIEW_LIMIT} chars]`
      : serialized
  } catch {
    return "[unserializable body]"
  }
}

function formatHeadersForLog(headers: Headers): Record<string, string> {
  return Object.fromEntries(Array.from(headers.entries()))
}

function extractGraphErrorInfo(
  url: string,
  method: string,
  requestBody: string | undefined,
  status: number,
  responseBody: string,
): GraphRequestErrorInfo {
  let requestId: string | undefined
  let clientRequestId: string | undefined

  try {
    const parsed = JSON.parse(responseBody)
    requestId = parsed?.error?.innerError?.["request-id"]
    clientRequestId = parsed?.error?.innerError?.["client-request-id"]
  } catch {}

  return {
    url,
    method,
    status,
    responseBody,
    requestBody,
    requestId,
    clientRequestId,
    message: `HTTP error! status: ${status}, body: ${responseBody}`,
  }
}

function isRecurrencePatchDateError(info: GraphRequestErrorInfo | null): boolean {
  if (!info) return false

  return (
    info.method === "PATCH" &&
    info.status === 400 &&
    info.responseBody?.includes("recurrence.range.startDate") === true &&
    info.responseBody?.includes("Microsoft.OData.Edm.Date") === true
  )
}

function isFutureOrCurrentDateTime(value: DateTimeTimeZone): boolean {
  if (!value.dateTime) return false

  const dueDateOnly = value.dateTime.slice(0, 10)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDateOnly)) return false

  const today = new Date()
  const todayDateOnly = `${today.getUTCFullYear()}-${String(today.getUTCMonth() + 1).padStart(2, "0")}-${String(
    today.getUTCDate(),
  ).padStart(2, "0")}`

  return dueDateOnly >= todayDateOnly
}

function parseDateOnly(value: string): Date | null {
  const dateOnly = value.slice(0, 10)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateOnly)) return null

  const [year, month, day] = dateOnly.split("-").map(Number)
  return new Date(Date.UTC(year, month - 1, day))
}

function formatDateOnly(date: Date): string {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`
}

function addDays(date: Date, days: number): Date {
  const next = new Date(date.getTime())
  next.setUTCDate(next.getUTCDate() + days)
  return next
}

function addMonths(date: Date, months: number): Date {
  const year = date.getUTCFullYear()
  const month = date.getUTCMonth()
  const day = date.getUTCDate()

  const targetMonthStart = new Date(Date.UTC(year, month + months, 1))
  const lastDayOfMonth = new Date(
    Date.UTC(targetMonthStart.getUTCFullYear(), targetMonthStart.getUTCMonth() + 1, 0),
  ).getUTCDate()
  return new Date(
    Date.UTC(targetMonthStart.getUTCFullYear(), targetMonthStart.getUTCMonth(), Math.min(day, lastDayOfMonth)),
  )
}

function diffDays(start: Date, end: Date): number {
  return Math.floor((end.getTime() - start.getTime()) / 86400000)
}

function diffMonths(start: Date, end: Date): number {
  return (end.getUTCFullYear() - start.getUTCFullYear()) * 12 + (end.getUTCMonth() - start.getUTCMonth())
}

function diffYears(start: Date, end: Date): number {
  return end.getUTCFullYear() - start.getUTCFullYear()
}

function getWeekdayIndex(day: string): number {
  const normalized = day.toLowerCase()
  const mapping: Record<string, number> = {
    sunday: 0,
    monday: 1,
    tuesday: 2,
    wednesday: 3,
    thursday: 4,
    friday: 5,
    saturday: 6,
  }

  return mapping[normalized] ?? -1
}

function startOfWeek(date: Date, firstDayOfWeek: string): Date {
  const first = getWeekdayIndex(firstDayOfWeek || "sunday")
  const offset = (date.getUTCDay() - first + 7) % 7
  return addDays(date, -offset)
}

function getMonthOccurrenceIndex(date: Date): number {
  return Math.floor((date.getUTCDate() - 1) / 7)
}

function getLastOccurrenceIndexInMonth(date: Date): number {
  const nextWeek = addDays(date, 7)
  return nextWeek.getUTCMonth() !== date.getUTCMonth() ? getMonthOccurrenceIndex(date) : -1
}

function matchesRelativePattern(date: Date, daysOfWeek: string[], index?: string): boolean {
  const allowedDays = daysOfWeek.map(getWeekdayIndex)
  if (!allowedDays.includes(date.getUTCDay())) return false

  const occurrenceIndex = getMonthOccurrenceIndex(date)
  if (!index || index === "first") return occurrenceIndex === 0
  if (index === "second") return occurrenceIndex === 1
  if (index === "third") return occurrenceIndex === 2
  if (index === "fourth") return occurrenceIndex === 3
  if (index === "last") return getLastOccurrenceIndexInMonth(date) !== -1

  return false
}

function matchesRecurrenceDate(date: Date, recurrence: PatternedRecurrence, anchorDate: Date): boolean {
  const interval = recurrence.pattern.interval || 1

  switch (recurrence.pattern.type) {
    case "daily":
      return diffDays(anchorDate, date) % interval === 0
    case "weekly": {
      const daysOfWeek = recurrence.pattern.daysOfWeek || []
      if (daysOfWeek.length === 0) return false
      const weekOffset = Math.floor(
        diffDays(
          startOfWeek(anchorDate, recurrence.pattern.firstDayOfWeek || "sunday"),
          startOfWeek(date, recurrence.pattern.firstDayOfWeek || "sunday"),
        ) / 7,
      )
      return weekOffset % interval === 0 && daysOfWeek.map(getWeekdayIndex).includes(date.getUTCDay())
    }
    case "absoluteMonthly":
      return diffMonths(anchorDate, date) % interval === 0 && recurrence.pattern.dayOfMonth === date.getUTCDate()
    case "relativeMonthly":
      return (
        diffMonths(anchorDate, date) % interval === 0 &&
        matchesRelativePattern(date, recurrence.pattern.daysOfWeek || [], recurrence.pattern.index)
      )
    case "absoluteYearly":
      return (
        diffYears(anchorDate, date) % interval === 0 &&
        recurrence.pattern.month === date.getUTCMonth() + 1 &&
        recurrence.pattern.dayOfMonth === date.getUTCDate()
      )
    case "relativeYearly":
      return (
        diffYears(anchorDate, date) % interval === 0 &&
        recurrence.pattern.month === date.getUTCMonth() + 1 &&
        matchesRelativePattern(date, recurrence.pattern.daysOfWeek || [], recurrence.pattern.index)
      )
    default:
      return false
  }
}

function findNextCurrentOccurrence(task: Task): Date | null {
  if (!task.recurrence || !task.dueDateTime?.dateTime) return null

  const anchorDate = parseDateOnly(task.recurrence.range.startDate || task.dueDateTime.dateTime)
  if (!anchorDate) return null

  const today = parseDateOnly(new Date().toISOString())
  if (!today) return null

  const rangeEndDate =
    task.recurrence.range.type === "endDate" && task.recurrence.range.endDate
      ? parseDateOnly(task.recurrence.range.endDate)
      : null

  let occurrenceCount = 0
  for (let offset = 0; offset <= 3660; offset++) {
    const candidate = addDays(anchorDate, offset)
    if (rangeEndDate && candidate > rangeEndDate) return null

    if (matchesRecurrenceDate(candidate, task.recurrence, anchorDate)) {
      occurrenceCount++
      if (
        task.recurrence.range.type === "numbered" &&
        task.recurrence.range.numberOfOccurrences &&
        occurrenceCount > task.recurrence.range.numberOfOccurrences
      ) {
        return null
      }

      if (candidate >= today) {
        return candidate
      }
    }
  }

  return null
}

function replaceDatePortion(dateTime: string, date: Date): string {
  const dateOnly = formatDateOnly(date)
  const timePortion = dateTime.includes("T") ? dateTime.slice(dateTime.indexOf("T")) : "T00:00:00.0000000"
  return `${dateOnly}${timePortion}`
}

function shiftDateTimeByDays(dateTime: string, days: number): string {
  const baseDate = parseDateOnly(dateTime)
  if (!baseDate) return dateTime
  return replaceDatePortion(dateTime, addDays(baseDate, days))
}

async function patchRecurringTaskDateFields(
  token: string,
  listId: string,
  taskId: string,
  existingTask: Task,
  taskBody: Record<string, unknown>,
): Promise<RecurringDatePatchResult> {
  if (!existingTask.recurrence) {
    return { error: `Task ${taskId} does not have recurrence to preserve while updating date fields.` }
  }

  const effectiveDueDate =
    taskBody.dueDateTime === undefined
      ? existingTask.dueDateTime
      : taskBody.dueDateTime === null
        ? null
        : (taskBody.dueDateTime as DateTimeTimeZone)

  if (!effectiveDueDate) {
    return {
      error:
        "Cannot update dueDateTime/reminderDateTime on a recurring task while ending up with no dueDateTime. " +
        "Provide a dueDateTime or clear recurrence explicitly.",
    }
  }

  const clearedRecurrence = await makeGraphRequest<Task>(
    `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
    token,
    "PATCH",
    {
      recurrence: null,
    },
  )

  if (!clearedRecurrence) {
    return { error: `Failed to temporarily clear recurrence for task ${taskId} before updating date fields.` }
  }

  const bodyWithoutRecurrence = { ...taskBody }
  delete bodyWithoutRecurrence.recurrence

  let intermediateResponse: Task | null = clearedRecurrence
  if (Object.keys(bodyWithoutRecurrence).length > 0) {
    intermediateResponse = await makeGraphRequest<Task>(
      `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
      token,
      "PATCH",
      bodyWithoutRecurrence,
    )

    if (!intermediateResponse) {
      return { error: `Failed to update task ${taskId} after clearing recurrence.` }
    }
  }

  const restoredRecurrence = await makeGraphRequest<Task>(
    `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
    token,
    "PATCH",
    {
      dueDateTime: effectiveDueDate,
      recurrence: buildRecurrencePatchPayloadFromExisting(existingTask.recurrence),
    },
  )

  if (!restoredRecurrence) {
    return { error: `Failed to restore recurrence for task ${taskId} after updating its date fields.` }
  }

  return { task: restoredRecurrence }
}

function formatDateTime(value?: DateTimeTimeZone | null): string | null {
  if (!value?.dateTime) return null

  const date = new Date(value.dateTime)
  if (Number.isNaN(date.getTime())) {
    return `${value.dateTime} (${value.timeZone})`
  }

  return `${date.toLocaleString()} (${value.timeZone})`
}

function formatRecurrence(recurrence?: PatternedRecurrence | null): string | null {
  if (!recurrence) return null

  const details = [`${recurrence.pattern.type} every ${recurrence.pattern.interval}`]
  const hasSentinelEndDate = recurrence.range.endDate === "0001-01-01"

  if (recurrence.pattern.daysOfWeek?.length) {
    details.push(`on ${recurrence.pattern.daysOfWeek.join(", ")}`)
  }

  if (recurrence.pattern.dayOfMonth) {
    details.push(`day ${recurrence.pattern.dayOfMonth}`)
  }

  if (recurrence.range.startDate) {
    details.push(`starting ${recurrence.range.startDate}`)
  }

  if (recurrence.range.type !== "noEnd" && recurrence.range.endDate && !hasSentinelEndDate) {
    details.push(`until ${recurrence.range.endDate}`)
  } else if (recurrence.range.numberOfOccurrences) {
    details.push(`for ${recurrence.range.numberOfOccurrences} occurrence(s)`)
  }

  return details.join(", ")
}

function formatTask(task: Task): string {
  let taskInfo = `ID: ${task.id}\nTitle: ${task.title}`

  if (task.status) {
    const status = task.status === "completed" ? "✓" : "○"
    taskInfo = `${status} ${taskInfo}`
    taskInfo += `\nStatus: ${task.status}`
  }

  const start = formatDateTime(task.startDateTime)
  const due = formatDateTime(task.dueDateTime)
  const reminder = formatDateTime(task.reminderDateTime)
  const completed = formatDateTime(task.completedDateTime)
  const recurrence = formatRecurrence(task.recurrence)

  if (start) taskInfo += `\nStart: ${start}`
  if (due) taskInfo += `\nDue: ${due}`
  if (reminder) taskInfo += `\nReminder: ${reminder}`
  if (completed) taskInfo += `\nCompleted: ${completed}`
  if (task.importance) taskInfo += `\nImportance: ${task.importance}`
  if (task.isReminderOn !== undefined) taskInfo += `\nReminder Enabled: ${task.isReminderOn ? "Yes" : "No"}`
  if (task.hasAttachments !== undefined) taskInfo += `\nHas Attachments: ${task.hasAttachments ? "Yes" : "No"}`
  if (recurrence) taskInfo += `\nRecurrence: ${recurrence}`
  if (task.categories && task.categories.length > 0) taskInfo += `\nCategories: ${task.categories.join(", ")}`

  if (task.linkedResources && task.linkedResources.length > 0) {
    const linkedSummary = task.linkedResources
      .map(
        (resource) =>
          resource.displayName || resource.applicationName || resource.webUrl || resource.id || "Linked item",
      )
      .join(", ")
    taskInfo += `\nLinked Resources: ${linkedSummary}`
  }

  if (task.createdDateTime) taskInfo += `\nCreated: ${new Date(task.createdDateTime).toLocaleString()}`
  if (task.lastModifiedDateTime) taskInfo += `\nLast Modified: ${new Date(task.lastModifiedDateTime).toLocaleString()}`
  if (task.bodyLastModifiedDateTime)
    taskInfo += `\nBody Modified: ${new Date(task.bodyLastModifiedDateTime).toLocaleString()}`

  if (task.body && task.body.content && task.body.content.trim() !== "") {
    taskInfo += `\nDescription: ${task.body.content}`
  }

  if (task.checklistItems && task.checklistItems.length > 0) {
    const checklistSummary = task.checklistItems
      .map((item) => `  ${item.isChecked ? "✓" : "○"} ${item.displayName}`)
      .join("\n")
    taskInfo += `\nChecklist (${task.checklistItems.filter((i) => i.isChecked).length}/${task.checklistItems.length}):\n${checklistSummary}`
  }

  return `${taskInfo}\n---`
}

const GRAPH_URL_PREFIX = "https://graph.microsoft.com/"

function isAllowedGraphUrl(url: string): boolean {
  return url.startsWith(GRAPH_URL_PREFIX)
}

function isTaskFileAttachment(value: unknown): value is TaskFileAttachment {
  return Boolean(
    value &&
      typeof value === "object" &&
      "id" in value &&
      "name" in value &&
      typeof (value as { id?: unknown }).id === "string" &&
      typeof (value as { name?: unknown }).name === "string",
  )
}

// === Tools ===

// Auth status
server.tool(
  "auth-status",
  "Check if you're authenticated with Microsoft Graph API. Shows current token status and expiration time.",
  {},
  async () => {
    const tokens = await tokenManager.getTokens()

    if (!tokens) {
      return {
        content: [
          {
            type: "text",
            text: "Not authenticated. Please run 'npx mstodo-setup' or 'pnpm run setup' to authenticate with Microsoft.",
          },
        ],
      }
    }

    const isExpired = Date.now() > tokens.expiresAt
    const expiryTime = new Date(tokens.expiresAt).toLocaleString()

    const isPersonal = await isPersonalMicrosoftAccount()
    let accountMessage = ""

    if (isPersonal) {
      accountMessage =
        "\n\nWARNING: You are using a personal Microsoft account. " +
        "Microsoft To Do API access is typically not available for personal accounts " +
        "through the Microsoft Graph API. You may encounter 'MailboxNotEnabledForRESTAPI' errors."
    }

    if (isExpired) {
      return {
        content: [
          {
            type: "text",
            text: `Authentication expired at ${expiryTime}. Will attempt to refresh when you call any API.${accountMessage}`,
          },
        ],
      }
    } else {
      return {
        content: [
          {
            type: "text",
            text: `Authenticated. Token expires at ${expiryTime}.${accountMessage}`,
          },
        ],
      }
    }
  },
)

server.tool(
  "refresh-auth-token",
  "Force a Microsoft Graph token refresh using the stored refresh token and report the new expiration time.",
  {},
  async () => {
    const previousTokens = await tokenManager.getTokens()

    if (!previousTokens) {
      return {
        content: [
          {
            type: "text",
            text: "Not authenticated. Please run 'npx mstodo-setup' or 'pnpm run setup' to authenticate with Microsoft.",
          },
        ],
      }
    }

    const refreshedTokens = await tokenManager.getTokens({ forceRefresh: true })

    if (!refreshedTokens) {
      return {
        content: [
          {
            type: "text",
            text:
              "Failed to refresh the Microsoft Graph token. Reauthentication may be required." +
              `\nToken file: ${tokenManager.getTokenFilePath()}`,
          },
        ],
      }
    }

    const refreshedExpiryTime = new Date(refreshedTokens.expiresAt).toLocaleString()
    const previousExpiryTime = new Date(previousTokens.expiresAt).toLocaleString()
    const didExpiryChange = refreshedTokens.expiresAt !== previousTokens.expiresAt

    return {
      content: [
        {
          type: "text",
          text: didExpiryChange
            ? `Authentication refreshed successfully. Previous expiry: ${previousExpiryTime}. New expiry: ${refreshedExpiryTime}.\nToken file: ${tokenManager.getTokenFilePath()}`
            : `Authentication is already current. Current expiry: ${refreshedExpiryTime}.\nToken file: ${tokenManager.getTokenFilePath()}`,
        },
      ],
    }
  },
)

// Task List tools
server.tool(
  "get-task-lists",
  "Get all Microsoft Todo task lists (the top-level containers that organize your tasks). Shows list names, IDs, and indicates default or shared lists.",
  {},
  async () => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)

      if (!response) {
        return {
          content: [{ type: "text", text: "Failed to retrieve task lists" }],
        }
      }

      const lists = response.value || []
      if (lists.length === 0) {
        return {
          content: [{ type: "text", text: "No task lists found." }],
        }
      }

      const formattedLists = lists.map((list) => {
        let wellKnownInfo = ""
        if (list.wellknownListName && list.wellknownListName !== "none") {
          if (list.wellknownListName === "defaultList") {
            wellKnownInfo = " (Default Tasks List)"
          } else if (list.wellknownListName === "flaggedEmails") {
            wellKnownInfo = " (Flagged Emails)"
          }
        }

        let sharingInfo = ""
        if (list.isShared) {
          sharingInfo = list.isOwner ? " (Shared by you)" : " (Shared with you)"
        }

        return `ID: ${list.id}\nName: ${list.displayName}${wellKnownInfo}${sharingInfo}\n---`
      })

      return {
        content: [{ type: "text", text: `Your task lists:\n\n${formattedLists.join("\n")}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching task lists: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-task-list",
  "Get a single Microsoft Todo task list by ID.",
  {
    listId: z.string().describe("ID of the task list"),
    select: z.string().optional().describe("Comma-separated list of properties to include"),
  },
  async ({ listId, select }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const queryParams = new URLSearchParams()
      if (select) queryParams.append("$select", select)

      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}${queryString ? "?" + queryString : ""}`
      const list = await makeGraphRequest<TaskList>(url, token)

      if (!list) {
        return {
          content: [{ type: "text", text: `Failed to retrieve task list: ${listId}` }],
        }
      }

      const metadata = []
      if (list.wellknownListName && list.wellknownListName !== "none") metadata.push(`Type: ${list.wellknownListName}`)
      if (list.isShared !== undefined) metadata.push(`Shared: ${list.isShared ? "Yes" : "No"}`)
      if (list.isOwner !== undefined) metadata.push(`Owner: ${list.isOwner ? "Yes" : "No"}`)

      return {
        content: [
          {
            type: "text",
            text: `Task list details:\n\nID: ${list.id}\nName: ${list.displayName}${metadata.length ? `\n${metadata.join("\n")}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching task list: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-task-lists-delta",
  "Track changes to Microsoft Todo task lists using the Graph delta API.",
  {
    deltaUrl: z
      .string()
      .optional()
      .describe("Full @odata.nextLink or @odata.deltaLink URL from a previous delta response"),
    deltaToken: z.string().optional().describe("Delta token from a previous delta response"),
    skipToken: z.string().optional().describe("Skip token from a previous delta response"),
    select: z.string().optional().describe("Comma-separated list of properties to include on the initial request"),
    maxPageSize: z.number().int().min(1).optional().describe("Preferred maximum number of lists returned"),
  },
  async ({ deltaUrl, deltaToken, skipToken, select, maxPageSize }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      let url = deltaUrl || `${MS_GRAPH_BASE}/me/todo/lists/delta`
      if (deltaUrl && !isAllowedGraphUrl(deltaUrl)) {
        return {
          content: [
            {
              type: "text",
              text: "Invalid deltaUrl: must be a Microsoft Graph API URL (https://graph.microsoft.com/...)",
            },
          ],
        }
      }
      if (!deltaUrl) {
        const queryParams = new URLSearchParams()
        if (deltaToken) queryParams.append("$deltatoken", deltaToken)
        if (skipToken) queryParams.append("$skiptoken", skipToken)
        if (select) queryParams.append("$select", select)
        const queryString = queryParams.toString()
        if (queryString) url += `?${queryString}`
      }

      const data = await makeGraphRequest<DeltaResponse<TaskList>>(url, token)
      if (!data) {
        return {
          content: [{ type: "text", text: "Failed to fetch task list delta" }],
        }
      }
      const formattedLists = (data.value || [])
        .map((list) => `ID: ${list.id}\nName: ${list.displayName}\n---`)
        .join("\n")

      return {
        content: [
          {
            type: "text",
            text:
              `Task list delta results:\n\n${formattedLists || "No changed lists returned."}` +
              `${data["@odata.nextLink"] ? `\n\nNext Link:\n${data["@odata.nextLink"]}` : ""}` +
              `${data["@odata.deltaLink"] ? `\n\nDelta Link:\n${data["@odata.deltaLink"]}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching task list delta: ${error}` }],
      }
    }
  },
)

server.tool(
  "create-task-list",
  "Create a new task list (top-level container) in Microsoft Todo to help organize your tasks into categories or projects.",
  {
    displayName: z.string().describe("Name of the new task list"),
  },
  async ({ displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists`, token, "POST", {
        displayName,
      })

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to create task list: ${displayName}` }],
        }
      }

      return {
        content: [
          { type: "text", text: `Task list created successfully!\nName: ${response.displayName}\nID: ${response.id}` },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error creating task list: ${error}` }],
      }
    }
  },
)

server.tool(
  "update-task-list",
  "Update the name of an existing task list (top-level container) in Microsoft Todo.",
  {
    listId: z.string().describe("ID of the task list to update"),
    displayName: z.string().describe("New name for the task list"),
  },
  async ({ listId, displayName }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<TaskList>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}`, token, "PATCH", {
        displayName,
      })

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to update task list with ID: ${listId}` }],
        }
      }

      return {
        content: [{ type: "text", text: `Task list updated successfully!\nNew name: ${response.displayName}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error updating task list: ${error}` }],
      }
    }
  },
)

server.tool(
  "delete-task-list",
  "Delete a task list (top-level container) from Microsoft Todo. This will remove the list and all tasks within it.",
  {
    listId: z.string().describe("ID of the task list to delete"),
  },
  async ({ listId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      await makeGraphRequest<null>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}`, token, "DELETE")

      return {
        content: [{ type: "text", text: `Task list with ID: ${listId} was successfully deleted.` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error deleting task list: ${error}` }],
      }
    }
  },
)

// Task tools
server.tool(
  "get-tasks",
  "Get tasks from a specific Microsoft Todo list. These are the main todo items that can contain checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    filter: z.string().optional().describe("OData $filter query (e.g., 'status eq \\'completed\\'')"),
    select: z.string().optional().describe("Comma-separated list of properties to include (e.g., 'id,title,status')"),
    orderby: z.string().optional().describe("Property to sort by (e.g., 'createdDateTime desc')"),
    top: z.number().optional().describe("Maximum number of tasks to retrieve"),
    skip: z.number().optional().describe("Number of tasks to skip"),
    count: z.boolean().optional().describe("Whether to include a count of tasks"),
  },
  async ({ listId, filter, select, orderby, top, skip, count }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const queryParams = new URLSearchParams()

      if (filter) queryParams.append("$filter", filter)
      if (select) queryParams.append("$select", select)
      if (orderby) queryParams.append("$orderby", orderby)
      if (top !== undefined) queryParams.append("$top", top.toString())
      if (skip !== undefined) queryParams.append("$skip", skip.toString())
      if (count !== undefined) queryParams.append("$count", count.toString())
      // Include navigation properties by default
      if (!queryParams.has("$expand")) queryParams.append("$expand", "linkedResources,checklistItems")

      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks${queryString ? "?" + queryString : ""}`

      const response = await makeGraphRequest<{ value: Task[]; "@odata.count"?: number }>(url, token)

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to retrieve tasks for list: ${listId}` }],
        }
      }

      const tasks = response.value || []
      if (tasks.length === 0) {
        return {
          content: [{ type: "text", text: `No tasks found in list with ID: ${listId}` }],
        }
      }

      const formattedTasks = tasks.map((task) => formatTask(task))

      let countInfo = ""
      if (count && response["@odata.count"] !== undefined) {
        countInfo = `Total count: ${response["@odata.count"]}\n\n`
      }

      return {
        content: [{ type: "text", text: `Tasks in list ${listId}:\n\n${countInfo}${formattedTasks.join("\n")}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching tasks: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-task",
  "Get a single Microsoft Todo task by ID.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    select: z.string().optional().describe("Comma-separated list of properties to include"),
  },
  async ({ listId, taskId, select }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const queryParams = new URLSearchParams()
      if (select) queryParams.append("$select", select)
      // Include navigation properties by default
      queryParams.append("$expand", "linkedResources,checklistItems")

      const queryString = queryParams.toString()
      const url = `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}${queryString ? "?" + queryString : ""}`
      const task = await makeGraphRequest<Task>(url, token)

      if (!task) {
        return {
          content: [{ type: "text", text: `Failed to retrieve task: ${taskId}` }],
        }
      }

      return {
        content: [{ type: "text", text: `Task details:\n\n${formatTask(task)}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching task: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-tasks-delta",
  "Track changes to tasks in a Microsoft Todo list using the Graph delta API.",
  {
    listId: z.string().describe("ID of the task list"),
    deltaUrl: z
      .string()
      .optional()
      .describe("Full @odata.nextLink or @odata.deltaLink URL from a previous delta response"),
    deltaToken: z.string().optional().describe("Delta token from a previous delta response"),
    skipToken: z.string().optional().describe("Skip token from a previous delta response"),
    select: z.string().optional().describe("Comma-separated list of properties to include on the initial request"),
    top: z.number().int().min(1).optional().describe("Maximum number of tasks to retrieve on the initial request"),
    expand: z.string().optional().describe("OData $expand expression for the initial request"),
    maxPageSize: z.number().int().min(1).optional().describe("Preferred maximum number of tasks returned"),
  },
  async ({ listId, deltaUrl, deltaToken, skipToken, select, top, expand, maxPageSize }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      let url = deltaUrl || `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/delta`
      if (deltaUrl && !isAllowedGraphUrl(deltaUrl)) {
        return {
          content: [
            {
              type: "text",
              text: "Invalid deltaUrl: must be a Microsoft Graph API URL (https://graph.microsoft.com/...)",
            },
          ],
        }
      }
      if (!deltaUrl) {
        const queryParams = new URLSearchParams()
        if (deltaToken) queryParams.append("$deltatoken", deltaToken)
        if (skipToken) queryParams.append("$skiptoken", skipToken)
        if (select) queryParams.append("$select", select)
        if (top !== undefined) queryParams.append("$top", top.toString())
        if (expand) queryParams.append("$expand", expand)
        const queryString = queryParams.toString()
        if (queryString) url += `?${queryString}`
      }

      const data = await makeGraphRequest<DeltaResponse<Task>>(url, token)
      if (!data) {
        return {
          content: [{ type: "text", text: `Failed to fetch task delta for list: ${listId}` }],
        }
      }
      const formattedTasks = (data.value || []).map((task) => formatTask(task)).join("\n")

      return {
        content: [
          {
            type: "text",
            text:
              `Task delta results for list ${listId}:\n\n${formattedTasks || "No changed tasks returned."}` +
              `${data["@odata.nextLink"] ? `\n\nNext Link:\n${data["@odata.nextLink"]}` : ""}` +
              `${data["@odata.deltaLink"] ? `\n\nDelta Link:\n${data["@odata.deltaLink"]}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching task delta: ${error}` }],
      }
    }
  },
)

server.tool(
  "create-task",
  "Create a new task in a specific Microsoft Todo list. A task is the main todo item that can have a title, description, due date, and other properties.",
  {
    listId: z.string().describe("ID of the task list"),
    title: z.string().describe("Title of the task"),
    body: z.string().optional().describe("Description or body content of the task"),
    dueDateTime: z.string().optional().describe("Due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    completedDateTime: z.string().optional().describe("Completion date in ISO format"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("Task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("Reminder date and time in ISO format"),
    recurrence: recurrenceSchema.optional().describe("Structured recurrence definition"),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("Status of the task"),
    categories: z.array(z.string()).optional().describe("Categories associated with the task"),
    linkedResources: z.array(linkedResourceSchema).optional().describe("Linked resources to create with the task"),
  },
  async ({
    listId,
    title,
    body,
    dueDateTime,
    completedDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    recurrence,
    status,
    categories,
    linkedResources,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const taskBody: any = { title }

      if (body) {
        taskBody.body = { content: body, contentType: "text" }
      }

      if (dueDateTime) taskBody.dueDateTime = buildDateTimeTimeZone(dueDateTime)
      if (completedDateTime) taskBody.completedDateTime = buildDateTimeTimeZone(completedDateTime)
      if (importance) taskBody.importance = importance

      if (isReminderOn !== undefined) taskBody.isReminderOn = isReminderOn
      if (reminderDateTime) taskBody.reminderDateTime = buildDateTimeTimeZone(reminderDateTime)
      if (isReminderOn === false) taskBody.reminderDateTime = null

      if (recurrence) {
        if (!taskBody.dueDateTime) {
          return {
            content: [
              {
                type: "text",
                text: "Recurring tasks require dueDateTime. Please provide dueDateTime or remove the recurrence.",
              },
            ],
          }
        }
        taskBody.recurrence = buildRecurrencePayload(recurrence)
      }

      if (status) taskBody.status = status
      if (categories && categories.length > 0) taskBody.categories = categories
      if (linkedResources && linkedResources.length > 0) taskBody.linkedResources = linkedResources

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token,
        "POST",
        taskBody,
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to create task in list: ${listId}` }],
        }
      }

      return {
        content: [{ type: "text", text: `Task created successfully!\nID: ${response.id}\nTitle: ${response.title}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error creating task: ${error}` }],
      }
    }
  },
)

server.tool(
  "update-task",
  "Update an existing task in Microsoft Todo. Allows changing any properties of the task including title, due date, importance, etc.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to update"),
    title: z.string().optional().describe("New title of the task"),
    body: z.string().optional().describe("New description or body content of the task"),
    dueDateTime: z.string().optional().describe("New due date in ISO format (e.g., 2023-12-31T23:59:59Z)"),
    completedDateTime: z
      .string()
      .optional()
      .describe("New completion date in ISO format. Pass an empty string to clear it."),
    importance: z.enum(["low", "normal", "high"]).optional().describe("New task importance"),
    isReminderOn: z.boolean().optional().describe("Whether to enable reminder for this task"),
    reminderDateTime: z.string().optional().describe("New reminder date and time in ISO format"),
    recurrence: recurrenceSchema.nullable().optional().describe("New recurrence definition. Pass null to clear it."),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("New status of the task"),
    categories: z.array(z.string()).optional().describe("New categories associated with the task"),
  },
  async ({
    listId,
    taskId,
    title,
    body,
    dueDateTime,
    completedDateTime,
    importance,
    isReminderOn,
    reminderDateTime,
    recurrence,
    status,
    categories,
  }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      let existingTask: Task | null = null
      const taskBody: any = {}

      if (title !== undefined) taskBody.title = title

      if (body !== undefined) {
        taskBody.body = { content: body, contentType: "text" }
      }

      if (dueDateTime !== undefined) {
        taskBody.dueDateTime = dueDateTime === "" ? null : buildDateTimeTimeZone(dueDateTime)
      }

      if (completedDateTime !== undefined) {
        taskBody.completedDateTime = completedDateTime === "" ? null : buildDateTimeTimeZone(completedDateTime)
      }

      if (importance !== undefined) taskBody.importance = importance
      if (isReminderOn !== undefined) taskBody.isReminderOn = isReminderOn

      if (reminderDateTime !== undefined) {
        taskBody.reminderDateTime = reminderDateTime === "" ? null : buildDateTimeTimeZone(reminderDateTime)
      }

      if (isReminderOn === false) taskBody.reminderDateTime = null

      if (recurrence !== undefined) {
        if (recurrence === null) {
          taskBody.recurrence = null
        } else {
          if (taskBody.dueDateTime === undefined) {
            existingTask = await makeGraphRequest<Task>(
              `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
              token,
            )

            if (!existingTask) {
              return {
                content: [
                  {
                    type: "text",
                    text: `Failed to load task ${taskId} to determine the due date required for recurrence updates.`,
                  },
                ],
              }
            }

            if (!existingTask.dueDateTime) {
              return {
                content: [
                  {
                    type: "text",
                    text: "Microsoft Graph requires dueDateTime when adding or updating recurrence. This task has no due date, so include dueDateTime in the update request.",
                  },
                ],
              }
            }

            if (!isFutureOrCurrentDateTime(existingTask.dueDateTime)) {
              return {
                content: [
                  {
                    type: "text",
                    text: "Microsoft Graph requires dueDateTime when adding or updating recurrence. This task's existing due date is already in the past, so specify a new dueDateTime in the update request.",
                  },
                ],
              }
            }

            taskBody.dueDateTime = existingTask.dueDateTime
          } else if (taskBody.dueDateTime === null) {
            return {
              content: [
                {
                  type: "text",
                  text: "Cannot clear dueDateTime while setting recurrence. Microsoft Graph requires dueDateTime for recurrence updates.",
                },
              ],
            }
          }

          taskBody.recurrence = buildRecurrencePatchPayload(recurrence)
        }
      }

      if (status !== undefined) taskBody.status = status
      if (categories !== undefined) taskBody.categories = categories

      const isRecurringTaskDateAdjustment =
        recurrence === undefined && (taskBody.dueDateTime !== undefined || taskBody.reminderDateTime !== undefined)

      if (isRecurringTaskDateAdjustment) {
        existingTask =
          existingTask ||
          (await makeGraphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token))

        if (!existingTask) {
          return {
            content: [{ type: "text", text: `Failed to load task ${taskId} before updating recurring task dates.` }],
          }
        }

        if (existingTask.recurrence) {
          const recurringPatchResult = await patchRecurringTaskDateFields(token, listId, taskId, existingTask, taskBody)
          if (recurringPatchResult.error || !recurringPatchResult.task) {
            return {
              content: [
                { type: "text", text: recurringPatchResult.error || `Failed to update recurring task ${taskId}.` },
              ],
            }
          }

          return {
            content: [
              {
                type: "text",
                text: `Task updated successfully!\nID: ${recurringPatchResult.task.id}\nTitle: ${recurringPatchResult.task.title}`,
              },
            ],
          }
        }
      }

      if (Object.keys(taskBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify at least one property to change.",
            },
          ],
        }
      }

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
        "PATCH",
        taskBody,
      )

      if (!response) {
        if (taskBody.recurrence !== undefined && isRecurrencePatchDateError(lastGraphRequestError)) {
          return {
            content: [
              {
                type: "text",
                text: "Microsoft Graph rejected the recurrence update. This appears to be a Graph To Do API limitation affecting PATCH updates to recurrence.range.startDate.",
              },
            ],
          }
        }

        return {
          content: [{ type: "text", text: `Failed to update task with ID: ${taskId} in list: ${listId}` }],
        }
      }

      return {
        content: [{ type: "text", text: `Task updated successfully!\nID: ${response.id}\nTitle: ${response.title}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error updating task: ${error}` }],
      }
    }
  },
)

server.tool(
  "delete-task",
  "Delete a task from a Microsoft Todo list. This will remove the task and all its checklist items (subtasks).",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to delete"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      await makeGraphRequest<null>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token, "DELETE")

      return {
        content: [{ type: "text", text: `Task with ID: ${taskId} was successfully deleted from list: ${listId}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error deleting task: ${error}` }],
      }
    }
  },
)

server.tool(
  "move-task",
  "Move a task from one list to another, preserving checklist items and most metadata. Tasks with attachments cannot be moved. Creation timestamps cannot be preserved due to API limitations.",
  {
    sourceListId: z.string().describe("ID of the source task list"),
    sourceTaskId: z.string().describe("ID of the task to move"),
    targetListId: z.string().describe("ID of the target task list"),
  },
  async ({ sourceListId, sourceTaskId, targetListId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      // Get the original task with all details
      const originalTask = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${sourceTaskId}`,
        token,
      )

      if (!originalTask) {
        return {
          content: [{ type: "text", text: `Failed to retrieve task: ${sourceTaskId}` }],
        }
      }

      // Reject if task has attachments
      if (originalTask.hasAttachments === true) {
        return {
          content: [
            {
              type: "text",
              text: `Cannot move task "${originalTask.title}" because it has attachments. Tasks with attachments cannot be moved between lists.`,
            },
          ],
        }
      }

      // Get checklist items
      const checklistResponse = await makeGraphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${sourceTaskId}/checklistItems`,
        token,
      )
      const checklistItems = checklistResponse?.value || []

      // Create the new task in the target list with all metadata
      const newTaskBody: any = { title: originalTask.title }

      if (originalTask.body) {
        newTaskBody.body = { content: originalTask.body.content, contentType: originalTask.body.contentType || "text" }
      }
      if (originalTask.dueDateTime) newTaskBody.dueDateTime = originalTask.dueDateTime
      if (originalTask.startDateTime) newTaskBody.startDateTime = originalTask.startDateTime
      if (originalTask.importance) newTaskBody.importance = originalTask.importance
      if (originalTask.isReminderOn !== undefined) newTaskBody.isReminderOn = originalTask.isReminderOn
      if (originalTask.reminderDateTime) newTaskBody.reminderDateTime = originalTask.reminderDateTime
      if (originalTask.status) newTaskBody.status = originalTask.status
      if (originalTask.categories) newTaskBody.categories = originalTask.categories
      if (originalTask.recurrence) newTaskBody.recurrence = originalTask.recurrence
      if (originalTask.linkedResources) newTaskBody.linkedResources = originalTask.linkedResources
      if (originalTask.completedDateTime) newTaskBody.completedDateTime = originalTask.completedDateTime

      const newTask = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`,
        token,
        "POST",
        newTaskBody,
      )

      if (!newTask) {
        return {
          content: [{ type: "text", text: "Failed to create task in target list" }],
        }
      }

      // Copy checklist items and track failures
      let checklistCopied = 0
      let checklistFailed = 0
      for (const item of checklistItems) {
        const result = await makeGraphRequest(
          `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks/${newTask.id}/checklistItems`,
          token,
          "POST",
          { displayName: item.displayName, isChecked: item.isChecked },
        )
        if (result) {
          checklistCopied++
        } else {
          checklistFailed++
        }
      }

      // Only delete the original if the new task was created and all checklist items copied
      let deleteStatus = "Skipped (checklist copy had failures)"
      if (checklistFailed === 0) {
        try {
          await makeGraphRequest(
            `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${sourceTaskId}`,
            token,
            "DELETE",
          )
          deleteStatus = "Yes"
        } catch {
          deleteStatus = "Failed"
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Successfully moved task "${originalTask.title}"\n\n` +
              `New Task ID: ${newTask.id}\n` +
              `Checklist items copied: ${checklistCopied}/${checklistItems.length}` +
              (checklistFailed > 0 ? ` (${checklistFailed} failed)` : "") +
              `\nOriginal task deleted: ${deleteStatus}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error moving task: ${error}` }],
      }
    }
  },
)

server.tool(
  "skip-task-to-current",
  "Advance a recurring task to the next occurrence on or after today. Preserves the recurrence and shifts the reminder by the same number of days when present.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the recurring task"),
    dryRun: z.boolean().optional().describe("Preview the calculated due/reminder dates without modifying the task"),
  },
  async ({ listId, taskId, dryRun = false }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const task = await makeGraphRequest<Task>(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token)
      if (!task) {
        return {
          content: [{ type: "text", text: `Failed to load task ${taskId}.` }],
        }
      }

      if (!task.recurrence) {
        return {
          content: [{ type: "text", text: "This task does not have recurrence, so there is nothing to skip." }],
        }
      }

      if (!task.dueDateTime?.dateTime) {
        return {
          content: [
            { type: "text", text: "Recurring tasks need dueDateTime in order to skip to the current occurrence." },
          ],
        }
      }

      const currentDueDate = parseDateOnly(task.dueDateTime.dateTime)
      const targetDueDate = findNextCurrentOccurrence(task)
      if (!currentDueDate || !targetDueDate) {
        return {
          content: [{ type: "text", text: "Could not calculate the next current occurrence for this recurring task." }],
        }
      }

      if (targetDueDate <= currentDueDate) {
        return {
          content: [
            {
              type: "text",
              text: `Task is already on the current occurrence.\nDue Date: ${formatDateOnly(currentDueDate)}`,
            },
          ],
        }
      }

      const deltaDays = diffDays(currentDueDate, targetDueDate)
      const nextDueDateTime: DateTimeTimeZone = {
        dateTime: replaceDatePortion(task.dueDateTime.dateTime, targetDueDate),
        timeZone: task.dueDateTime.timeZone,
      }

      const patchBody: Record<string, unknown> = { dueDateTime: nextDueDateTime }

      if (task.reminderDateTime?.dateTime) {
        patchBody.reminderDateTime = {
          dateTime: shiftDateTimeByDays(task.reminderDateTime.dateTime, deltaDays),
          timeZone: task.reminderDateTime.timeZone,
        }
      }

      if (dryRun) {
        let preview = `Skip Preview\nTask: ${task.title}\nCurrent Due Date: ${formatDateOnly(currentDueDate)}\nNew Due Date: ${formatDateOnly(targetDueDate)}`

        if (task.reminderDateTime?.dateTime) {
          preview +=
            `\nCurrent Reminder: ${task.reminderDateTime.dateTime} (${task.reminderDateTime.timeZone})` +
            `\nNew Reminder: ${(patchBody.reminderDateTime as DateTimeTimeZone).dateTime} ` +
            `(${(patchBody.reminderDateTime as DateTimeTimeZone).timeZone})`
        }

        preview += `\nShift Applied: ${deltaDays} day(s)`

        return {
          content: [{ type: "text", text: preview }],
        }
      }

      const recurringPatchResult = await patchRecurringTaskDateFields(token, listId, taskId, task, patchBody)
      if (recurringPatchResult.error || !recurringPatchResult.task) {
        return {
          content: [
            {
              type: "text",
              text: recurringPatchResult.error || `Failed to skip recurring task ${taskId} to the current occurrence.`,
            },
          ],
        }
      }

      let result =
        `Task advanced to current occurrence.\nID: ${recurringPatchResult.task.id}\nTitle: ${recurringPatchResult.task.title}` +
        `\nDue Date: ${formatDateOnly(targetDueDate)}\nShift Applied: ${deltaDays} day(s)`

      if (patchBody.reminderDateTime) {
        const newReminder = patchBody.reminderDateTime as DateTimeTimeZone
        result += `\nReminder: ${newReminder.dateTime} (${newReminder.timeZone})`
      }

      return {
        content: [{ type: "text", text: result }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error skipping recurring task to current occurrence: ${error}` }],
      }
    }
  },
)

// Linked resource tools
server.tool(
  "get-linked-resources",
  "Get linked resources for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<{ value: LinkedResource[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources`,
        token,
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to retrieve linked resources for task: ${taskId}` }],
        }
      }

      const linkedResources = response.value || []
      if (linkedResources.length === 0) {
        return {
          content: [{ type: "text", text: `No linked resources found for task: ${taskId}` }],
        }
      }

      const formattedResources = linkedResources.map((resource) => {
        const details = [
          `ID: ${resource.id || "Unknown"}`,
          `Display Name: ${resource.displayName || "Unknown"}`,
          `Application: ${resource.applicationName || "Unknown"}`,
        ]

        if (resource.webUrl) details.push(`URL: ${resource.webUrl}`)
        if (resource.externalId) details.push(`External ID: ${resource.externalId}`)

        return `${details.join("\n")}\n---`
      })

      return {
        content: [{ type: "text", text: `Linked resources for task ${taskId}:\n\n${formattedResources.join("\n")}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching linked resources: ${error}` }],
      }
    }
  },
)

server.tool(
  "create-linked-resource",
  "Create a linked resource for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    webUrl: z.string().optional().describe("Deep link to the linked item"),
    applicationName: z.string().optional().describe("Source application name"),
    displayName: z.string().optional().describe("Display name for the linked item"),
    externalId: z.string().optional().describe("External identifier from the source system"),
  },
  async ({ listId, taskId, webUrl, applicationName, displayName, externalId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<LinkedResource>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources`,
        token,
        "POST",
        { webUrl, applicationName, displayName, externalId },
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to create linked resource for task: ${taskId}` }],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Linked resource created successfully!\nID: ${response.id || "Unknown"}` +
              `${response.displayName ? `\nDisplay Name: ${response.displayName}` : ""}` +
              `${response.webUrl ? `\nURL: ${response.webUrl}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error creating linked resource: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-linked-resource",
  "Get a single linked resource by ID for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    linkedResourceId: z.string().describe("ID of the linked resource"),
  },
  async ({ listId, taskId, linkedResourceId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const resource = await makeGraphRequest<LinkedResource>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources/${linkedResourceId}`,
        token,
      )

      if (!resource) {
        return {
          content: [{ type: "text", text: `Failed to retrieve linked resource: ${linkedResourceId}` }],
        }
      }

      const details = [`ID: ${resource.id || "Unknown"}`, `Display Name: ${resource.displayName || "Unknown"}`]
      if (resource.applicationName) details.push(`Application: ${resource.applicationName}`)
      if (resource.webUrl) details.push(`URL: ${resource.webUrl}`)
      if (resource.externalId) details.push(`External ID: ${resource.externalId}`)

      return {
        content: [{ type: "text", text: `Linked resource details:\n\n${details.join("\n")}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching linked resource: ${error}` }],
      }
    }
  },
)

server.tool(
  "update-linked-resource",
  "Update an existing linked resource on a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    linkedResourceId: z.string().describe("ID of the linked resource to update"),
    webUrl: z.string().optional().describe("New deep link URL"),
    applicationName: z.string().optional().describe("New source application name"),
    displayName: z.string().optional().describe("New display name"),
    externalId: z.string().optional().describe("New external identifier"),
  },
  async ({ listId, taskId, linkedResourceId, webUrl, applicationName, displayName, externalId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const body: any = {}
      if (webUrl !== undefined) body.webUrl = webUrl
      if (applicationName !== undefined) body.applicationName = applicationName
      if (displayName !== undefined) body.displayName = displayName
      if (externalId !== undefined) body.externalId = externalId

      if (Object.keys(body).length === 0) {
        return {
          content: [{ type: "text", text: "No properties provided for update." }],
        }
      }

      const response = await makeGraphRequest<LinkedResource>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources/${linkedResourceId}`,
        token,
        "PATCH",
        body,
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to update linked resource: ${linkedResourceId}` }],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Linked resource updated successfully!\nID: ${response.id || "Unknown"}` +
              `${response.displayName ? `\nDisplay Name: ${response.displayName}` : ""}` +
              `${response.webUrl ? `\nURL: ${response.webUrl}` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error updating linked resource: ${error}` }],
      }
    }
  },
)

server.tool(
  "delete-linked-resource",
  "Delete a linked resource from a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    linkedResourceId: z.string().describe("ID of the linked resource to delete"),
  },
  async ({ listId, taskId, linkedResourceId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      await makeGraphRequest<null>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/linkedResources/${linkedResourceId}`,
        token,
        "DELETE",
      )

      return {
        content: [{ type: "text", text: `Linked resource ${linkedResourceId} deleted from task: ${taskId}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error deleting linked resource: ${error}` }],
      }
    }
  },
)

// Attachment tools
server.tool(
  "get-attachments",
  "List file attachments for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<{ value: TaskFileAttachment[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments`,
        token,
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to retrieve attachments for task: ${taskId}` }],
        }
      }

      const attachments = response.value || []
      if (attachments.length === 0) {
        return {
          content: [{ type: "text", text: `No attachments found for task: ${taskId}` }],
        }
      }

      const formattedAttachments = attachments.map((attachment) => {
        const details = [
          `ID: ${attachment.id}`,
          `Name: ${attachment.name}`,
          `Size: ${attachment.size ?? "Unknown"} bytes`,
        ]

        if (attachment.contentType) details.push(`Content Type: ${attachment.contentType}`)
        if (attachment.lastModifiedDateTime) {
          details.push(`Last Modified: ${new Date(attachment.lastModifiedDateTime).toLocaleString()}`)
        }

        return `${details.join("\n")}\n---`
      })

      return {
        content: [{ type: "text", text: `Attachments for task ${taskId}:\n\n${formattedAttachments.join("\n")}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching attachments: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-attachment",
  "Get a single file attachment for a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    attachmentId: z.string().describe("ID of the attachment"),
  },
  async ({ listId, taskId, attachmentId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<{ value?: TaskFileAttachment } | TaskFileAttachment>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`,
        token,
      )

      let attachment: TaskFileAttachment | undefined
      if (response && typeof response === "object" && "value" in response) {
        attachment = response.value
      } else if (isTaskFileAttachment(response)) {
        attachment = response
      }

      if (!attachment) {
        return {
          content: [{ type: "text", text: `Failed to retrieve attachment: ${attachmentId}` }],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Attachment details:\n\nID: ${attachment.id}\nName: ${attachment.name}` +
              `${attachment.contentType ? `\nContent Type: ${attachment.contentType}` : ""}` +
              `${attachment.size !== undefined ? `\nSize: ${attachment.size} bytes` : ""}` +
              `${attachment.lastModifiedDateTime ? `\nLast Modified: ${new Date(attachment.lastModifiedDateTime).toLocaleString()}` : ""}` +
              `${attachment.contentBytes ? `\nContent Bytes Present: Yes (${attachment.contentBytes.length} base64 chars)` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching attachment: ${error}` }],
      }
    }
  },
)

server.tool(
  "create-attachment",
  "Create a small file attachment on a Microsoft Todo task. For files larger than 3 MB, use create-attachment-upload-session.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    name: z.string().describe("Attachment display name"),
    contentBytes: z.string().describe("Base64-encoded file contents"),
    contentType: z.string().optional().describe("Attachment content type"),
  },
  async ({ listId, taskId, name, contentBytes, contentType }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<TaskFileAttachment>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments`,
        token,
        "POST",
        {
          "@odata.type": "#microsoft.graph.taskFileAttachment",
          name,
          contentBytes,
          contentType,
        },
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to create attachment for task: ${taskId}` }],
        }
      }

      return {
        content: [
          {
            type: "text",
            text:
              `Attachment created successfully!\nID: ${response.id}\nName: ${response.name}` +
              `${response.size !== undefined ? `\nSize: ${response.size} bytes` : ""}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error creating attachment: ${error}` }],
      }
    }
  },
)

server.tool(
  "create-attachment-upload-session",
  "Create an upload session for a large file attachment on a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    name: z.string().describe("Attachment display name"),
    size: z.number().int().min(0).describe("Attachment size in bytes"),
  },
  async ({ listId, taskId, name, size }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<UploadSession>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/createUploadSession`,
        token,
        "POST",
        {
          attachmentInfo: {
            attachmentType: "file",
            name,
            size,
          },
        },
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to create attachment upload session for task: ${taskId}` }],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Attachment upload session created successfully!\nUpload URL: ${response.uploadUrl}\nExpiration: ${response.expirationDateTime}\nNext Expected Ranges: ${response.nextExpectedRanges.join(", ")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error creating attachment upload session: ${error}` }],
      }
    }
  },
)

server.tool(
  "delete-attachment",
  "Delete a file attachment from a Microsoft Todo task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    attachmentId: z.string().describe("ID of the attachment"),
  },
  async ({ listId, taskId, attachmentId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      await makeGraphRequest<null>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/attachments/${attachmentId}`,
        token,
        "DELETE",
      )

      return {
        content: [
          { type: "text", text: `Attachment with ID: ${attachmentId} was successfully deleted from task: ${taskId}` },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error deleting attachment: ${error}` }],
      }
    }
  },
)

// Checklist item tools
server.tool(
  "get-checklist-items",
  "Get checklist items (subtasks) for a specific task. Checklist items are smaller steps or components that belong to a parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const taskResponse = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
      )

      const taskTitle = taskResponse ? taskResponse.title : "Unknown Task"

      const response = await makeGraphRequest<{ value: ChecklistItem[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to retrieve checklist items for task: ${taskId}` }],
        }
      }

      const items = response.value || []
      if (items.length === 0) {
        return {
          content: [{ type: "text", text: `No checklist items found for task "${taskTitle}" (ID: ${taskId})` }],
        }
      }

      const formattedItems = items.map((item) => {
        const status = item.isChecked ? "✓" : "○"
        let itemInfo = `${status} ${item.displayName} (ID: ${item.id})`

        if (item.createdDateTime) {
          itemInfo += `\nCreated: ${new Date(item.createdDateTime).toLocaleString()}`
        }

        if (item.checkedDateTime) {
          itemInfo += `\nChecked: ${new Date(item.checkedDateTime).toLocaleString()}`
        }

        return itemInfo
      })

      return {
        content: [
          {
            type: "text",
            text: `Checklist items for task "${taskTitle}" (ID: ${taskId}):\n\n${formattedItems.join("\n\n")}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching checklist items: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-checklist-item",
  "Get a single checklist item (subtask) by ID.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item"),
  },
  async ({ listId, taskId, checklistItemId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const item = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
        token,
      )

      if (!item) {
        return {
          content: [{ type: "text", text: `Failed to retrieve checklist item: ${checklistItemId}` }],
        }
      }

      const status = item.isChecked ? "Checked" : "Not checked"
      let details = `${status}: ${item.displayName}\nID: ${item.id}`
      if (item.createdDateTime) details += `\nCreated: ${new Date(item.createdDateTime).toLocaleString()}`
      if (item.checkedDateTime) details += `\nChecked: ${new Date(item.checkedDateTime).toLocaleString()}`

      return {
        content: [{ type: "text", text: `Checklist item details:\n\n${details}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching checklist item: ${error}` }],
      }
    }
  },
)

server.tool(
  "create-checklist-item",
  "Create a new checklist item (subtask) for a task. Checklist items help break down a task into smaller, manageable steps.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    displayName: z.string().describe("Text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
    checkedDateTime: z.string().optional().describe("Completion timestamp in ISO format"),
    createdDateTime: z.string().optional().describe("Creation timestamp in ISO format"),
  },
  async ({ listId, taskId, displayName, isChecked, checkedDateTime, createdDateTime }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const requestBody: any = { displayName }
      if (isChecked !== undefined) requestBody.isChecked = isChecked
      if (checkedDateTime !== undefined) requestBody.checkedDateTime = checkedDateTime
      if (createdDateTime !== undefined) requestBody.createdDateTime = createdDateTime

      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems`,
        token,
        "POST",
        requestBody,
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to create checklist item for task: ${taskId}` }],
        }
      }

      return {
        content: [
          {
            type: "text",
            text: `Checklist item created successfully!\nContent: ${response.displayName}\nID: ${response.id}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error creating checklist item: ${error}` }],
      }
    }
  },
)

server.tool(
  "update-checklist-item",
  "Update an existing checklist item (subtask). Allows changing the text content or completion status of the subtask.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to update"),
    displayName: z.string().optional().describe("New text content of the checklist item"),
    isChecked: z.boolean().optional().describe("Whether the item is checked off"),
    checkedDateTime: z
      .string()
      .optional()
      .describe("Completion timestamp in ISO format. Pass an empty string to clear it."),
    createdDateTime: z.string().optional().describe("Creation timestamp in ISO format"),
  },
  async ({ listId, taskId, checklistItemId, displayName, isChecked, checkedDateTime, createdDateTime }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const requestBody: any = {}
      if (displayName !== undefined) requestBody.displayName = displayName
      if (isChecked !== undefined) requestBody.isChecked = isChecked
      if (checkedDateTime !== undefined) requestBody.checkedDateTime = checkedDateTime === "" ? null : checkedDateTime
      if (createdDateTime !== undefined) requestBody.createdDateTime = createdDateTime

      if (Object.keys(requestBody).length === 0) {
        return {
          content: [
            {
              type: "text",
              text: "No properties provided for update. Please specify at least one checklist item property to change.",
            },
          ],
        }
      }

      const response = await makeGraphRequest<ChecklistItem>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
        token,
        "PATCH",
        requestBody,
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to update checklist item with ID: ${checklistItemId}` }],
        }
      }

      const statusText = response.isChecked ? "Checked" : "Not checked"

      return {
        content: [
          {
            type: "text",
            text: `Checklist item updated successfully!\nContent: ${response.displayName}\nStatus: ${statusText}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error updating checklist item: ${error}` }],
      }
    }
  },
)

server.tool(
  "delete-checklist-item",
  "Delete a checklist item (subtask) from a task. This removes just the specific subtask, not the parent task.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task"),
    checklistItemId: z.string().describe("ID of the checklist item to delete"),
  },
  async ({ listId, taskId, checklistItemId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      await makeGraphRequest<null>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}/checklistItems/${checklistItemId}`,
        token,
        "DELETE",
      )

      return {
        content: [
          {
            type: "text",
            text: `Checklist item with ID: ${checklistItemId} was successfully deleted from task: ${taskId}`,
          },
        ],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error deleting checklist item: ${error}` }],
      }
    }
  },
)

// Convenience tools
server.tool(
  "complete-task",
  "Mark a task as completed. A shortcut for update-task with status 'completed'.",
  {
    listId: z.string().describe("ID of the task list"),
    taskId: z.string().describe("ID of the task to complete"),
  },
  async ({ listId, taskId }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const response = await makeGraphRequest<Task>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`,
        token,
        "PATCH",
        { status: "completed" },
      )

      if (!response) {
        return {
          content: [{ type: "text", text: `Failed to complete task: ${taskId}` }],
        }
      }

      return {
        content: [{ type: "text", text: `Task completed!\nTitle: ${response.title}\nID: ${response.id}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error completing task: ${error}` }],
      }
    }
  },
)

server.tool(
  "search-tasks",
  "Search for tasks across all lists by keyword, status, importance, or due date range. Returns matching tasks with their list names.",
  {
    keyword: z.string().optional().describe("Search keyword to match in task titles (case-insensitive)"),
    status: z
      .enum(["notStarted", "inProgress", "completed", "waitingOnOthers", "deferred"])
      .optional()
      .describe("Filter by task status"),
    importance: z.enum(["low", "normal", "high"]).optional().describe("Filter by importance"),
    dueBefore: z.string().optional().describe("Filter tasks due before this date (ISO format, e.g. 2024-12-31)"),
    dueAfter: z.string().optional().describe("Filter tasks due after this date (ISO format, e.g. 2024-01-01)"),
    includeCompleted: z.boolean().optional().describe("Include completed tasks in results (default: false)"),
  },
  async ({ keyword, status, importance, dueBefore, dueAfter, includeCompleted = false }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      // Get all lists
      const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)
      if (!listsResponse) {
        return {
          content: [{ type: "text", text: "Failed to retrieve task lists" }],
        }
      }

      const allResults: { listName: string; task: Task }[] = []

      for (const list of listsResponse.value || []) {
        // Build OData filter
        const filters: string[] = []
        if (status) filters.push(`status eq '${status}'`)
        if (!includeCompleted && !status) filters.push("status ne 'completed'")
        if (importance) filters.push(`importance eq '${importance}'`)

        const queryParams = new URLSearchParams()
        if (filters.length > 0) queryParams.append("$filter", filters.join(" and "))
        queryParams.append("$expand", "linkedResources,checklistItems")

        const queryString = queryParams.toString()
        const url = `${MS_GRAPH_BASE}/me/todo/lists/${list.id}/tasks${queryString ? "?" + queryString : ""}`

        const tasksResponse = await makeGraphRequest<{ value: Task[] }>(url, token)
        if (!tasksResponse) continue

        for (const task of tasksResponse.value || []) {
          // Client-side keyword filter (OData doesn't support contains on title for To Do)
          if (keyword && !task.title.toLowerCase().includes(keyword.toLowerCase())) continue

          // Client-side date range filters
          if (dueBefore && task.dueDateTime) {
            const dueDate = task.dueDateTime.dateTime.slice(0, 10)
            if (dueDate > dueBefore.slice(0, 10)) continue
          }
          if (dueAfter && task.dueDateTime) {
            const dueDate = task.dueDateTime.dateTime.slice(0, 10)
            if (dueDate < dueAfter.slice(0, 10)) continue
          }
          // If filtering by date range, skip tasks without due dates
          if ((dueBefore || dueAfter) && !task.dueDateTime) continue

          allResults.push({ listName: list.displayName, task })
        }
      }

      if (allResults.length === 0) {
        return {
          content: [{ type: "text", text: "No tasks found matching your criteria." }],
        }
      }

      const formatted = allResults.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")

      return {
        content: [{ type: "text", text: `Found ${allResults.length} task(s):\n\n${formatted}` }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error searching tasks: ${error}` }],
      }
    }
  },
)

server.tool(
  "get-todays-tasks",
  "Get all tasks due today or overdue across all lists. Provides a unified daily view of what needs attention.",
  {
    includeOverdue: z.boolean().optional().describe("Include overdue tasks (default: true)"),
    includeNoDueDate: z.boolean().optional().describe("Include tasks with no due date (default: false)"),
  },
  async ({ includeOverdue = true, includeNoDueDate = false }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const today = new Date()
      const todayStr = `${today.getUTCFullYear()}-${String(today.getUTCMonth() + 1).padStart(2, "0")}-${String(today.getUTCDate()).padStart(2, "0")}`

      const listsResponse = await makeGraphRequest<{ value: TaskList[] }>(`${MS_GRAPH_BASE}/me/todo/lists`, token)
      if (!listsResponse) {
        return {
          content: [{ type: "text", text: "Failed to retrieve task lists" }],
        }
      }

      const todayTasks: { listName: string; task: Task }[] = []
      const overdueTasks: { listName: string; task: Task }[] = []
      const noDueDateTasks: { listName: string; task: Task }[] = []

      for (const list of listsResponse.value || []) {
        const queryParams = new URLSearchParams()
        queryParams.append("$filter", "status ne 'completed'")
        queryParams.append("$expand", "linkedResources,checklistItems")

        const url = `${MS_GRAPH_BASE}/me/todo/lists/${list.id}/tasks?${queryParams.toString()}`
        const tasksResponse = await makeGraphRequest<{ value: Task[] }>(url, token)
        if (!tasksResponse) continue

        for (const task of tasksResponse.value || []) {
          if (!task.dueDateTime) {
            if (includeNoDueDate) {
              noDueDateTasks.push({ listName: list.displayName, task })
            }
            continue
          }

          const dueDate = task.dueDateTime.dateTime.slice(0, 10)

          if (dueDate === todayStr) {
            todayTasks.push({ listName: list.displayName, task })
          } else if (dueDate < todayStr && includeOverdue) {
            overdueTasks.push({ listName: list.displayName, task })
          }
        }
      }

      let output = ""

      if (overdueTasks.length > 0) {
        output += `OVERDUE (${overdueTasks.length}):\n\n`
        output += overdueTasks.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")
        output += "\n\n"
      }

      if (todayTasks.length > 0) {
        output += `DUE TODAY (${todayTasks.length}):\n\n`
        output += todayTasks.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")
        output += "\n\n"
      }

      if (noDueDateTasks.length > 0) {
        output += `NO DUE DATE (${noDueDateTasks.length}):\n\n`
        output += noDueDateTasks.map(({ listName, task }) => `[${listName}]\n${formatTask(task)}`).join("\n")
        output += "\n\n"
      }

      if (!output) {
        output = "No tasks due today and no overdue tasks. You're all caught up!"
      } else {
        output = `Daily Task Summary (${todayStr})\n${"=".repeat(40)}\n\n${output}`
      }

      return {
        content: [{ type: "text", text: output.trim() }],
      }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error fetching today's tasks: ${error}` }],
      }
    }
  },
)

// Bulk operations
server.tool(
  "archive-completed-tasks",
  "Move completed tasks older than a specified number of days from one list to another (archive) list. Useful for cleaning up active lists while preserving historical tasks.",
  {
    sourceListId: z.string().describe("ID of the source list to archive tasks from"),
    targetListId: z.string().describe("ID of the target archive list"),
    olderThanDays: z
      .number()
      .min(0)
      .default(90)
      .describe("Archive tasks completed more than this many days ago (default: 90)"),
    dryRun: z
      .boolean()
      .optional()
      .default(false)
      .describe("If true, only preview what would be archived without making changes"),
  },
  async ({ sourceListId, targetListId, olderThanDays, dryRun }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const cutoffDate = new Date()
      cutoffDate.setDate(cutoffDate.getDate() - olderThanDays)

      const tasksResponse = await makeGraphRequest<{ value: Task[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks?$filter=status eq 'completed'`,
        token,
      )

      if (!tasksResponse || !tasksResponse.value) {
        return {
          content: [{ type: "text", text: "Failed to retrieve tasks from source list" }],
        }
      }

      const tasksToArchive = tasksResponse.value.filter((task) => {
        if (!task.completedDateTime?.dateTime) return false
        const completedDate = new Date(task.completedDateTime.dateTime)
        return completedDate < cutoffDate
      })

      if (tasksToArchive.length === 0) {
        return {
          content: [{ type: "text", text: `No completed tasks found older than ${olderThanDays} days.` }],
        }
      }

      if (dryRun) {
        let preview = `Archive Preview\n`
        preview += `Would archive ${tasksToArchive.length} tasks completed before ${cutoffDate.toLocaleDateString()}\n\n`

        tasksToArchive.forEach((task) => {
          const completedDate = task.completedDateTime?.dateTime
            ? new Date(task.completedDateTime.dateTime).toLocaleDateString()
            : "Unknown"
          preview += `- ${task.title} (completed: ${completedDate})\n`
        })

        return { content: [{ type: "text", text: preview }] }
      }

      let successCount = 0
      let failedCopy: string[] = []
      let failedDelete: string[] = []

      for (const task of tasksToArchive) {
        try {
          const createResponse = await makeGraphRequest(
            `${MS_GRAPH_BASE}/me/todo/lists/${targetListId}/tasks`,
            token,
            "POST",
            {
              title: task.title,
              status: "completed",
              body: task.body,
              importance: task.importance,
              completedDateTime: task.completedDateTime,
              dueDateTime: task.dueDateTime,
              startDateTime: task.startDateTime,
              reminderDateTime: task.reminderDateTime,
              isReminderOn: task.isReminderOn,
              recurrence: task.recurrence,
              categories: task.categories,
              linkedResources: task.linkedResources,
            },
          )

          if (!createResponse) {
            failedCopy.push(task.title)
            continue
          }

          // Only delete source after successful copy
          try {
            await makeGraphRequest(`${MS_GRAPH_BASE}/me/todo/lists/${sourceListId}/tasks/${task.id}`, token, "DELETE")
            successCount++
          } catch {
            failedDelete.push(task.title)
          }
        } catch (error) {
          failedCopy.push(task.title)
        }
      }

      let result = `Archive Complete\n`
      result += `Successfully archived ${successCount} of ${tasksToArchive.length} tasks\n`
      result += `Tasks completed before ${cutoffDate.toLocaleDateString()} were moved.\n`

      if (failedCopy.length > 0) {
        result += `\nFailed to copy ${failedCopy.length} tasks:\n`
        failedCopy.forEach((title) => {
          result += `- ${title}\n`
        })
      }

      if (failedDelete.length > 0) {
        result += `\nCopied but failed to delete source (duplicates exist in both lists):\n`
        failedDelete.forEach((title) => {
          result += `- ${title}\n`
        })
      }

      return { content: [{ type: "text", text: result }] }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error archiving tasks: ${error}` }],
      }
    }
  },
)

server.tool(
  "reorganize-list",
  "Reorganize flat tasks in a list into category tasks with checklist items (subtasks). Takes a grouping spec and creates category parent tasks, moves matching task titles as checklist items under each category, and optionally deletes the originals. Includes dry-run for preview and idempotency checks to prevent duplicates.",
  {
    listId: z.string().describe("ID of the list to reorganize"),
    categories: z
      .array(
        z.object({
          title: z.string().describe("Category task title (e.g. 'Planning (2 weeks out)')"),
          taskTitles: z
            .array(z.string())
            .describe(
              "Titles of existing tasks to move under this category as checklist items. Matched case-insensitively.",
            ),
        }),
      )
      .describe("Category groupings - each becomes a parent task with checklist items"),
    deleteOriginals: z
      .boolean()
      .default(true)
      .describe("Delete original flat tasks after reorganizing (default: true)"),
    dryRun: z.boolean().default(false).describe("Preview changes without executing (default: false)"),
  },
  async ({ listId, categories, deleteOriginals, dryRun }) => {
    try {
      const token = await getAccessToken()
      if (!token) {
        return {
          content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
        }
      }

      const tasksResponse = await makeGraphRequest<{ value: Task[] }>(
        `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
        token,
      )

      if (!tasksResponse || !tasksResponse.value) {
        return {
          content: [{ type: "text", text: "Failed to retrieve tasks from list" }],
        }
      }

      const existingTasks = tasksResponse.value

      // Idempotency check
      const existingTitlesLower = new Set(existingTasks.map((t) => t.title.toLowerCase()))
      const alreadyExists = categories.filter((c) => existingTitlesLower.has(c.title.toLowerCase()))
      if (alreadyExists.length > 0) {
        return {
          content: [
            {
              type: "text",
              text:
                `Idempotency check failed. These category tasks already exist in the list:\n` +
                alreadyExists.map((c) => `- ${c.title}`).join("\n") +
                `\n\nDelete them first or rename your categories to avoid duplicates.`,
            },
          ],
        }
      }

      // Match task titles
      const tasksByTitleLower = new Map<string, Task>()
      for (const task of existingTasks) {
        tasksByTitleLower.set(task.title.toLowerCase(), task)
      }

      const matchedTaskIds: string[] = []
      const unmatchedTitles: string[] = []
      const categoryMatches: { category: string; matched: string[]; unmatched: string[] }[] = []

      for (const cat of categories) {
        const matched: string[] = []
        const unmatched: string[] = []

        for (const title of cat.taskTitles) {
          const task = tasksByTitleLower.get(title.toLowerCase())
          if (task) {
            matched.push(task.title)
            matchedTaskIds.push(task.id)
          } else {
            unmatched.push(title)
            unmatchedTitles.push(title)
          }
        }

        categoryMatches.push({ category: cat.title, matched, unmatched })
      }

      // Dry run
      if (dryRun) {
        let preview = `Reorganize Preview\n\n`
        preview += `List has ${existingTasks.length} tasks. Will create ${categories.length} category tasks.\n\n`

        for (const cm of categoryMatches) {
          preview += `Category: ${cm.category}\n`
          for (const m of cm.matched) {
            preview += `  + ${m}\n`
          }
          for (const u of cm.unmatched) {
            preview += `  ? ${u} (NOT FOUND)\n`
          }
        }

        if (deleteOriginals) {
          preview += `\nWill delete ${matchedTaskIds.length} original tasks after reorganizing.`
        }

        if (unmatchedTitles.length > 0) {
          preview += `\n\n${unmatchedTitles.length} task title(s) did not match any existing tasks.`
        }

        return { content: [{ type: "text", text: preview }] }
      }

      // Create category tasks + checklist items
      let createdCategories = 0
      let createdItems = 0
      let failedItems: string[] = []
      const successfullyMovedTaskIds = new Set<string>()

      for (const cat of categories) {
        try {
          const newTask = await makeGraphRequest<Task>(
            `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks`,
            token,
            "POST",
            { title: cat.title },
          )

          if (!newTask || !newTask.id) {
            failedItems.push(`Category: ${cat.title}`)
            continue
          }

          createdCategories++

          for (const title of cat.taskTitles) {
            const sourceTask = tasksByTitleLower.get(title.toLowerCase())
            const displayName = sourceTask ? sourceTask.title : title

            try {
              await makeGraphRequest(
                `${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${newTask.id}/checklistItems`,
                token,
                "POST",
                { displayName },
              )
              createdItems++
              if (sourceTask) {
                successfullyMovedTaskIds.add(sourceTask.id)
              }
            } catch (error) {
              failedItems.push(`Item: ${displayName}`)
            }
          }
        } catch (error) {
          failedItems.push(`Category: ${cat.title}`)
        }
      }

      // Only delete originals that were successfully added as checklist items
      let deletedCount = 0
      let deleteFailures: string[] = []

      if (deleteOriginals) {
        for (const taskId of matchedTaskIds) {
          if (!successfullyMovedTaskIds.has(taskId)) continue
          try {
            await makeGraphRequest(`${MS_GRAPH_BASE}/me/todo/lists/${listId}/tasks/${taskId}`, token, "DELETE")
            deletedCount++
          } catch (error) {
            const task = existingTasks.find((t) => t.id === taskId)
            deleteFailures.push(task?.title || taskId)
          }
        }
      }

      // Report
      let result = `Reorganize Complete\n\n`
      result += `Created ${createdCategories} category tasks with ${createdItems} checklist items.\n`

      if (deleteOriginals) {
        result += `Deleted ${deletedCount} original tasks.\n`
      }

      if (failedItems.length > 0) {
        result += `\nFailed to create:\n`
        failedItems.forEach((item) => {
          result += `- ${item}\n`
        })
      }

      if (deleteFailures.length > 0) {
        result += `\nFailed to delete:\n`
        deleteFailures.forEach((title) => {
          result += `- ${title}\n`
        })
      }

      if (unmatchedTitles.length > 0) {
        result += `\nUnmatched task titles (not found in list):\n`
        unmatchedTitles.forEach((title) => {
          result += `- ${title}\n`
        })
      }

      return { content: [{ type: "text", text: result }] }
    } catch (error) {
      return {
        content: [{ type: "text", text: `Error reorganizing list: ${error}` }],
      }
    }
  },
)

// Main function to start the server
export async function startServer(config?: ServerConfig): Promise<void> {
  try {
    if (config?.tokenFilePath) {
      process.env.MSTODO_TOKEN_FILE = config.tokenFilePath
      tokenManager.configure({ tokenFilePath: config.tokenFilePath })
    }

    if (config?.accessToken) {
      process.env.MS_TODO_ACCESS_TOKEN = config.accessToken
    }

    if (config?.refreshToken) {
      process.env.MS_TODO_REFRESH_TOKEN = config.refreshToken
    }

    // Check if using a personal Microsoft account and show warning if needed
    await isPersonalMicrosoftAccount()

    // Start the server
    const transport = new StdioServerTransport()
    await server.connect(transport)

    console.error("Server started and listening")
  } catch (error) {
    console.error("Error starting server:", error)
    throw error
  }
}

// Main entry point when executed directly
if (import.meta.url === `file://${process.argv[1]}`) {
  startServer().catch((error) => {
    console.error("Fatal error in main():", error)
    process.exit(1)
  })
}
