import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
import { z } from "zod"
import { makeGraphRequest, getAccessToken, MS_GRAPH_BASE } from "../graph-client.js"
import type { Task } from "../types.js"

export function register(server: McpServer) {
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
}
