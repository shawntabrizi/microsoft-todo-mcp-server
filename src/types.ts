import { z } from "zod"

export interface TaskList {
  id: string
  displayName: string
  isOwner?: boolean
  isShared?: boolean
  wellknownListName?: string
}

export interface DateTimeTimeZone {
  dateTime: string
  timeZone: string
}

export interface RecurrencePattern {
  type: string
  interval: number
  month?: number
  dayOfMonth?: number
  daysOfWeek?: string[]
  firstDayOfWeek?: string
  index?: string
}

export interface RecurrenceRange {
  type: string
  startDate: string
  endDate?: string
  recurrenceTimeZone?: string
  numberOfOccurrences?: number
}

export interface PatternedRecurrence {
  pattern: RecurrencePattern
  range: RecurrenceRange
}

export interface LinkedResource {
  id?: string
  webUrl?: string
  applicationName?: string
  displayName?: string
  externalId?: string
}

export interface TaskFileAttachment {
  id: string
  name: string
  contentType?: string
  size?: number
  lastModifiedDateTime?: string
  contentBytes?: string
}

export interface DeltaResponse<T> {
  value: T[]
  "@odata.count"?: number
  "@odata.nextLink"?: string
  "@odata.deltaLink"?: string
}

export interface Task {
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

export interface ChecklistItem {
  id: string
  displayName: string
  isChecked: boolean
  checkedDateTime?: string
  createdDateTime?: string
}

export interface UploadSession {
  uploadUrl: string
  expirationDateTime: string
  nextExpectedRanges: string[]
}

export interface ServerConfig {
  accessToken?: string
  refreshToken?: string
  tokenFilePath?: string
}

export type GraphRequestErrorInfo = {
  url: string
  method: string
  status?: number
  responseBody?: string
  requestBody?: string
  requestId?: string
  clientRequestId?: string
  message: string
}

export type RecurringDatePatchResult = {
  task?: Task
  error?: string
}

// === Zod schemas ===

export const recurrenceSchema = z.object({
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

export type RecurrenceInput = z.infer<typeof recurrenceSchema>

export const linkedResourceSchema = z.object({
  webUrl: z.string().optional().describe("Deep link to the linked item"),
  applicationName: z.string().optional().describe("Source application name"),
  displayName: z.string().optional().describe("Display title for the linked item"),
  externalId: z.string().optional().describe("External identifier from the source system"),
})
