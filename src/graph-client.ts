import { tokenManager } from "./token-manager.js"
import type { GraphRequestErrorInfo } from "./types.js"

export const MS_GRAPH_BASE = "https://graph.microsoft.com/v1.0"
const USER_AGENT = "microsoft-todo-mcp-server/1.0"
const GRAPH_URL_PREFIX = "https://graph.microsoft.com/"

export let lastGraphRequestError: GraphRequestErrorInfo | null = null

// Deduplication cache for non-idempotent requests (POST/PATCH/DELETE)
const inFlightRequests = new Map<string, Promise<any>>()

function getRequestKey(method: string, url: string, body?: string): string {
  return `${method}:${url}:${body ?? ""}`
}

export async function makeGraphRequest<T>(url: string, token: string, method = "GET", body?: any): Promise<T | null> {
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

      if (errorText.includes("MailboxNotEnabledForRESTAPI")) {
        console.error(
          "ERROR: MailboxNotEnabledForRESTAPI - " +
            "Microsoft To Do API is not available for personal Microsoft accounts.",
        )
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
    // Re-throw for DELETE so callers can distinguish failure from 204 success
    if (method === "DELETE") {
      throw error
    }
    return null
  }
}

export async function getAccessToken(forceRefresh = false): Promise<string | null> {
  try {
    const tokens = await tokenManager.getTokens({ forceRefresh })
    if (tokens) {
      return tokens.accessToken
    }
    console.error("No valid tokens available")
    return null
  } catch (error) {
    console.error("Error getting access token:", error)
    return null
  }
}

export async function isPersonalMicrosoftAccount(): Promise<boolean> {
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

    if (!response.ok) return false

    const userData = await response.json()
    const email = userData.mail || userData.userPrincipalName || ""
    const personalDomains = ["outlook.com", "hotmail.com", "live.com", "msn.com", "passport.com"]
    const domain = email.split("@")[1]?.toLowerCase()

    if (domain && personalDomains.some((d) => domain.includes(d))) {
      console.error(`WARNING: Personal Microsoft Account Detected (${email})`)
      return true
    }

    return false
  } catch (error) {
    console.error("Error checking account type:", error)
    return false
  }
}

export function isAllowedGraphUrl(url: string): boolean {
  return url.startsWith(GRAPH_URL_PREFIX)
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
