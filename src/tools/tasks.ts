import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import {
  makeGraphRequest,
  getAccessToken,
  MS_GRAPH_BASE,
  isAllowedGraphUrl,
  lastGraphRequestError,
} from "../graph-client.js"
import {
  formatTask,
  buildDateTimeTimeZone,
  buildRecurrencePayload,
  buildRecurrencePatchPayload,
  isFutureOrCurrentDateTime,
  isRecurrencePatchDateError,
  patchRecurringTaskDateFields,
  findNextCurrentOccurrence,
  parseDateOnly,
  formatDateOnly,
  replaceDatePortion,
  shiftDateTimeByDays,
  diffDays,
} from "../helpers.js"
import type { Task, DateTimeTimeZone, ChecklistItem, DeltaResponse } from "../types.js"
import { recurrenceSchema, linkedResourceSchema } from "../types.js"

export function register(server: McpServer) {
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
      expand: z
        .string()
        .optional()
        .describe(
          "Comma-separated navigation properties to expand (e.g., 'linkedResources,checklistItems'). Default: linkedResources",
        ),
    },
    async ({ listId, filter, select, orderby, top, skip, count, expand }) => {
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
        queryParams.append("$expand", expand || "linkedResources")

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
      expand: z
        .string()
        .optional()
        .describe(
          "Comma-separated navigation properties to expand (e.g., 'linkedResources,checklistItems'). Default: linkedResources,checklistItems",
        ),
    },
    async ({ listId, taskId, select, expand }) => {
      try {
        const token = await getAccessToken()
        if (!token) {
          return {
            content: [{ type: "text", text: "Failed to authenticate with Microsoft API" }],
          }
        }

        const queryParams = new URLSearchParams()
        if (select) queryParams.append("$select", select)
        queryParams.append("$expand", expand || "linkedResources,checklistItems")

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
            const recurringPatchResult = await patchRecurringTaskDateFields(
              token,
              listId,
              taskId,
              existingTask,
              taskBody,
            )
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
          newTaskBody.body = {
            content: originalTask.body.content,
            contentType: originalTask.body.contentType || "text",
          }
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
            content: [
              { type: "text", text: "Could not calculate the next current occurrence for this recurring task." },
            ],
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
                text:
                  recurringPatchResult.error || `Failed to skip recurring task ${taskId} to the current occurrence.`,
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
}
