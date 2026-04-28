import {
  IPublicClientApplication,
  InteractionRequiredAuthError,
} from "@azure/msal-browser";
import {
  AppConfig,
  ConversationTranscript,
  FeedbackItem,
  ODataResponse,
  TranscriptSchema,
} from "../types";
import { getDataverseScopes } from "../authConfig";
import { extractFeedback } from "./transcriptParser";

// Safety cap to prevent fetching excessive data in large environments
const MAX_TRANSCRIPTS = 5000;

/**
 * All known schema variants for the ConversationTranscript table.
 * Newer/core Dataverse environments use the unprefixed variant; older or
 * solution-based environments use the msdyn_ prefix.
 * The unprefixed variant is tried first since the user confirmed it.
 */
const SCHEMAS: TranscriptSchema[] = [
  {
    entitySetName: "conversationtranscripts",
    logicalName: "conversationtranscript",
    idField: "conversationtranscriptid",
    nameField: "name",
    contentField: "content",
    botAnnotationPrefixes: ["_botid_value", "_msdyn_botid_value"],
  },
  {
    entitySetName: "msdyn_conversationtranscripts",
    logicalName: "msdyn_conversationtranscript",
    idField: "msdyn_conversationtranscriptid",
    nameField: "msdyn_name",
    contentField: "msdyn_content",
    botAnnotationPrefixes: [
      "_msdyn_botid_value",
      "_msdyn_bot_value",
      "_msdyn_agentid_value",
    ],
  },
];

/** Returns the value if it is a non-empty string, otherwise undefined. */
function pickString(value: unknown): string | undefined {
  return typeof value === "string" && value.trim() ? value.trim() : undefined;
}

async function getAccessToken(
  instance: IPublicClientApplication,
  config: AppConfig
): Promise<string> {
  const accounts = instance.getAllAccounts();
  if (accounts.length === 0) throw new Error("Not authenticated. Please sign in.");

  const scopes = getDataverseScopes(config);
  try {
    const response = await instance.acquireTokenSilent({
      scopes,
      account: accounts[0],
    });
    return response.accessToken;
  } catch (error) {
    if (error instanceof InteractionRequiredAuthError) {
      const response = await instance.acquireTokenPopup({ scopes });
      return response.accessToken;
    }
    throw error;
  }
}

/**
 * Detects which API version (v9.2 / v9.1) the environment supports via WhoAmI.
 */
async function detectApiVersion(
  baseUrl: string,
  authHeader: string
): Promise<string> {
  const headers = { Authorization: authHeader, Accept: "application/json" };
  for (const version of ["v9.2", "v9.1", "v9.0"]) {
    try {
      const res = await fetch(`${baseUrl}/api/data/${version}/WhoAmI`, {
        headers,
      });
      if (res.ok || res.status === 403) return version;
    } catch {
      // try next version
    }
  }
  return "v9.2";
}

/**
 * Resolves the correct TranscriptSchema for this environment using three strategies:
 *
 * 1. Config override — if the user specified an entity set name manually, match it
 *    to a known schema (or synthesise a best-guess schema from the override name).
 * 2. EntityDefinitions metadata — query Dataverse for the registered entity set name
 *    for each known logical name.
 * 3. Direct probe — try a $top=1 request against each candidate entity set name and
 *    use the first one that responds with 2xx.
 */
async function resolveSchema(
  apiBase: string,
  authHeader: string,
  entitySetOverride: string | undefined
): Promise<TranscriptSchema> {
  const metaHeaders = {
    Authorization: authHeader,
    Accept: "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
  };

  // Strategy 1: manual override
  if (entitySetOverride) {
    const match = SCHEMAS.find(
      (s) => s.entitySetName.toLowerCase() === entitySetOverride.toLowerCase()
    );
    if (match) return match;
    // Unknown override — synthesise using the unprefixed field name convention
    // if the override looks unprefixed, otherwise msdyn_ convention.
    const base = SCHEMAS[entitySetOverride.startsWith("msdyn_") ? 1 : 0];
    return { ...base, entitySetName: entitySetOverride };
  }

  // Strategy 2: EntityDefinitions metadata
  for (const schema of SCHEMAS) {
    try {
      const url =
        `${apiBase}/EntityDefinitions` +
        `?$filter=LogicalName eq '${schema.logicalName}'&$select=EntitySetName`;
      const res = await fetch(url, { headers: metaHeaders });
      if (res.ok) {
        const data = (await res.json()) as ODataResponse<{
          EntitySetName?: string;
        }>;
        const entitySetName = pickString(data.value?.[0]?.EntitySetName);
        if (entitySetName) return { ...schema, entitySetName };
      }
    } catch {
      // try next schema
    }
  }

  // Strategy 3: direct probe with $top=1
  const probeHeaders = {
    ...metaHeaders,
    Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
  };
  for (const schema of SCHEMAS) {
    try {
      const url = `${apiBase}/${schema.entitySetName}?$select=${schema.idField}&$top=1`;
      const res = await fetch(url, { headers: probeHeaders });
      if (res.ok) return schema;
    } catch {
      // try next
    }
  }

  throw new Error(
    `Could not locate the ConversationTranscript table in your Dataverse environment.\n` +
      `Tried logical names: ${SCHEMAS.map((s) => s.logicalName).join(", ")}.\n\n` +
      `To fix this, go back to the setup screen (sign out → "Use a different environment"), ` +
      `expand Advanced options and enter the Entity set name exactly as shown in ` +
      `Power Apps maker portal → Tables → ConversationTranscript → Properties → Entity set name.`
  );
}

/**
 * Fetches all conversation transcripts from Dataverse and extracts feedback items.
 */
export async function fetchAllFeedback(
  instance: IPublicClientApplication,
  config: AppConfig
): Promise<FeedbackItem[]> {
  const token = await getAccessToken(instance, config);
  const baseUrl = config.dataverseUrl.replace(/\/$/, "");
  const authHeader = `Bearer ${token}`;

  // Detect API version, then resolve the correct table schema.
  const apiVersion = await detectApiVersion(baseUrl, authHeader);
  const apiBase = `${baseUrl}/api/data/${apiVersion}`;
  const schema = await resolveSchema(
    apiBase,
    authHeader,
    pickString(config.entitySetName)
  );

  const dataHeaders: HeadersInit = {
    Authorization: authHeader,
    Accept: "application/json",
    "OData-MaxVersion": "4.0",
    "OData-Version": "4.0",
    Prefer: "odata.include-annotations=OData.Community.Display.V1.FormattedValue",
  };

  const initialUrl =
    `${apiBase}/${schema.entitySetName}` +
    `?$orderby=createdon desc`;
    // Note: $select is intentionally omitted so that all lookup field annotations
    // (including the bot/agent display name FormattedValue) are returned by Dataverse.

  const allTranscripts: ConversationTranscript[] = [];
  let nextLink: string | null = initialUrl;

  while (nextLink !== null && allTranscripts.length < MAX_TRANSCRIPTS) {
    let response: Response;
    try {
      response = await fetch(nextLink, { headers: dataHeaders });
    } catch {
      throw new Error(
        `Network error reaching ${baseUrl}. ` +
          `Check the environment URL and that the Azure AD app redirect URI ` +
          `matches this page's origin (${window.location.origin}).`
      );
    }

    if (!response.ok) {
      const body = await response.text();
      if (response.status === 401 || response.status === 403) {
        throw new Error(
          `Access denied (${response.status}). ` +
            `Ensure the Azure AD app has Dynamics CRM → user_impersonation permission ` +
            `and the signed-in user has a Dataverse security role with read access to ConversationTranscript.`
        );
      }
      throw new Error(
        `Dataverse API error ${response.status} (${apiVersion}/${schema.entitySetName}): ${body}`
      );
    }

    const data = (await response.json()) as ODataResponse<ConversationTranscript>;
    allTranscripts.push(...(data.value ?? []));
    nextLink = data["@odata.nextLink"] ?? null;
  }

  const ANNOTATION_SUFFIX = "@OData.Community.Display.V1.FormattedValue";
  const feedbackItems: FeedbackItem[] = [];

  for (const transcript of allTranscripts) {
    const content = pickString(transcript[schema.contentField]);
    if (!content) continue;

    const id = pickString(transcript[schema.idField]) ?? "";

    // 1. Try schema-specific known bot annotation prefixes first.
    const agentFromKnownPrefix = schema.botAnnotationPrefixes
      .map((p) => pickString(transcript[`${p}${ANNOTATION_SUFFIX}`]))
      .find(Boolean);

    // 2. Scan ALL annotation keys for any that reference a bot/agent lookup.
    //    This handles environments with non-standard lookup field names.
    const agentFromScan = agentFromKnownPrefix
      ? undefined
      : Object.entries(transcript)
          .filter(
            ([k, v]) =>
              k.endsWith(ANNOTATION_SUFFIX) &&
              /bot|agent|chatbot/i.test(k) &&
              typeof v === "string" &&
              v.trim()
          )
          .map(([, v]) => v as string)[0];

    const agentName =
      agentFromKnownPrefix ??
      agentFromScan ??
      "Unknown Agent";

    const items = extractFeedback(id, content, agentName);
    feedbackItems.push(...items);
  }

  feedbackItems.sort((a, b) => b.timestamp.localeCompare(a.timestamp));
  return feedbackItems;
}

/**
 * Fetches the friendly display name of the Power Platform environment by
 * querying the Global Discovery Service, which returns the same name shown
 * in the Power Platform admin centre (e.g. "Global Azure Bootcamp").
 * Returns `undefined` if the request fails for any reason.
 */
export async function fetchEnvironmentName(
  instance: IPublicClientApplication,
  config: AppConfig
): Promise<string | undefined> {
  try {
    // The Discovery Service uses its own scope, separate from the Dataverse scope.
    const DISCO_SCOPE = "https://globaldisco.crm.dynamics.com/.default";
    const accounts = instance.getAllAccounts();
    if (accounts.length === 0) return undefined;

    let token: string;
    try {
      const res = await instance.acquireTokenSilent({
        scopes: [DISCO_SCOPE],
        account: accounts[0],
      });
      token = res.accessToken;
    } catch (e) {
      if (e instanceof InteractionRequiredAuthError) {
        const res = await instance.acquireTokenPopup({ scopes: [DISCO_SCOPE] });
        token = res.accessToken;
      } else {
        throw e;
      }
    }

    const res = await fetch(
      "https://globaldisco.crm.dynamics.com/api/discovery/v2.0/Instances",
      {
        headers: {
          Authorization: `Bearer ${token}`,
          Accept: "application/json",
        },
      }
    );

    if (!res.ok) {
      console.warn(
        `[fetchEnvironmentName] Discovery Service returned ${res.status}:`,
        await res.text()
      );
      return undefined;
    }

    const data = (await res.json()) as ODataResponse<{
      FriendlyName?: string;
      Url?: string;
      ApiUrl?: string;
    }>;

    // Normalise both sides: strip trailing slash, lower-case for comparison.
    const target = config.dataverseUrl.replace(/\/$/, "").toLowerCase();
    const match = data.value?.find(
      (env) =>
        env.Url?.replace(/\/$/, "").toLowerCase() === target ||
        env.ApiUrl?.replace(/\/$/, "").toLowerCase() === target
    );

    return pickString(match?.FriendlyName);
  } catch (err) {
    console.warn("[fetchEnvironmentName] failed:", err);
    return undefined;
  }
}

