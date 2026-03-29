import type {
  DateTimeTimeZone,
  PatternedRecurrence,
  RecurrenceInput,
  RecurringDatePatchResult,
  Task,
  TaskFileAttachment,
  GraphRequestErrorInfo,
} from "./types.js"
import { makeGraphRequest, lastGraphRequestError, MS_GRAPH_BASE } from "./graph-client.js"

// === Date/time builders ===

export function buildDateTimeTimeZone(dateTime: string): DateTimeTimeZone {
  return { dateTime, timeZone: "UTC" }
}

export function buildRecurrencePayload(recurrence: RecurrenceInput): PatternedRecurrence {
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

export function buildRecurrencePatchPayload(recurrence: RecurrenceInput): PatternedRecurrence {
  return {
    pattern: buildRecurrencePayload(recurrence).pattern,
    range: {} as PatternedRecurrence["range"],
  }
}

export function buildRecurrencePatchPayloadFromExisting(recurrence: PatternedRecurrence): PatternedRecurrence {
  return {
    pattern: recurrence.pattern,
    range: {} as PatternedRecurrence["range"],
  }
}

// === Date math ===

export function parseDateOnly(value: string): Date | null {
  const dateOnly = value.slice(0, 10)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dateOnly)) return null
  const [year, month, day] = dateOnly.split("-").map(Number)
  return new Date(Date.UTC(year, month - 1, day))
}

export function formatDateOnly(date: Date): string {
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth() + 1).padStart(2, "0")}-${String(date.getUTCDate()).padStart(2, "0")}`
}

export function addDays(date: Date, days: number): Date {
  const next = new Date(date.getTime())
  next.setUTCDate(next.getUTCDate() + days)
  return next
}

export function diffDays(start: Date, end: Date): number {
  return Math.floor((end.getTime() - start.getTime()) / 86400000)
}

function diffMonths(start: Date, end: Date): number {
  return (end.getUTCFullYear() - start.getUTCFullYear()) * 12 + (end.getUTCMonth() - start.getUTCMonth())
}

function diffYears(start: Date, end: Date): number {
  return end.getUTCFullYear() - start.getUTCFullYear()
}

function getWeekdayIndex(day: string): number {
  const mapping: Record<string, number> = {
    sunday: 0,
    monday: 1,
    tuesday: 2,
    wednesday: 3,
    thursday: 4,
    friday: 5,
    saturday: 6,
  }
  return mapping[day.toLowerCase()] ?? -1
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

export function findNextCurrentOccurrence(task: Task): Date | null {
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

export function replaceDatePortion(dateTime: string, date: Date): string {
  const dateOnly = formatDateOnly(date)
  const timePortion = dateTime.includes("T") ? dateTime.slice(dateTime.indexOf("T")) : "T00:00:00.0000000"
  return `${dateOnly}${timePortion}`
}

export function shiftDateTimeByDays(dateTime: string, days: number): string {
  const baseDate = parseDateOnly(dateTime)
  if (!baseDate) return dateTime
  return replaceDatePortion(dateTime, addDays(baseDate, days))
}

// === Recurrence patching ===

export function isRecurrencePatchDateError(info: GraphRequestErrorInfo | null): boolean {
  if (!info) return false
  return (
    info.method === "PATCH" &&
    info.status === 400 &&
    info.responseBody?.includes("recurrence.range.startDate") === true &&
    info.responseBody?.includes("Microsoft.OData.Edm.Date") === true
  )
}

export function isFutureOrCurrentDateTime(value: DateTimeTimeZone): boolean {
  if (!value.dateTime) return false
  const dueDateOnly = value.dateTime.slice(0, 10)
  if (!/^\d{4}-\d{2}-\d{2}$/.test(dueDateOnly)) return false

  const today = new Date()
  const todayDateOnly = `${today.getUTCFullYear()}-${String(today.getUTCMonth() + 1).padStart(2, "0")}-${String(today.getUTCDate()).padStart(2, "0")}`
  return dueDateOnly >= todayDateOnly
}

export async function patchRecurringTaskDateFields(
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
    { recurrence: null },
  )

  if (!clearedRecurrence) {
    return { error: `Failed to temporarily clear recurrence for task ${taskId} before updating date fields.` }
  }

  const bodyWithoutRecurrence = { ...taskBody }
  delete bodyWithoutRecurrence.recurrence

  if (Object.keys(bodyWithoutRecurrence).length > 0) {
    const intermediateResponse = await makeGraphRequest<Task>(
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

// === Formatting ===

export function formatDateTime(value?: DateTimeTimeZone | null): string | null {
  if (!value?.dateTime) return null
  const date = new Date(value.dateTime)
  if (Number.isNaN(date.getTime())) {
    return `${value.dateTime} (${value.timeZone})`
  }
  return `${date.toLocaleString()} (${value.timeZone})`
}

export function formatRecurrence(recurrence?: PatternedRecurrence | null): string | null {
  if (!recurrence) return null

  const details = [`${recurrence.pattern.type} every ${recurrence.pattern.interval}`]
  const hasSentinelEndDate = recurrence.range.endDate === "0001-01-01"

  if (recurrence.pattern.daysOfWeek?.length) details.push(`on ${recurrence.pattern.daysOfWeek.join(", ")}`)
  if (recurrence.pattern.dayOfMonth) details.push(`day ${recurrence.pattern.dayOfMonth}`)
  if (recurrence.range.startDate) details.push(`starting ${recurrence.range.startDate}`)

  if (recurrence.range.type !== "noEnd" && recurrence.range.endDate && !hasSentinelEndDate) {
    details.push(`until ${recurrence.range.endDate}`)
  } else if (recurrence.range.numberOfOccurrences) {
    details.push(`for ${recurrence.range.numberOfOccurrences} occurrence(s)`)
  }

  return details.join(", ")
}

export function formatTask(task: Task): string {
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

export function isTaskFileAttachment(value: unknown): value is TaskFileAttachment {
  return Boolean(
    value &&
      typeof value === "object" &&
      "id" in value &&
      "name" in value &&
      typeof (value as { id?: unknown }).id === "string" &&
      typeof (value as { name?: unknown }).name === "string",
  )
}
