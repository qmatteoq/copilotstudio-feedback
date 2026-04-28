import { Configuration } from "@azure/msal-browser";
import { AppConfig } from "./types";

const CONFIG_STORAGE_KEY = "copilotstudio_feedback_config";

export function getStoredConfig(): AppConfig | null {
  try {
    const stored = localStorage.getItem(CONFIG_STORAGE_KEY);
    if (!stored) {
      // Fall back to environment variables if set at build time
      const dataverseUrl = import.meta.env.VITE_DATAVERSE_URL as string | undefined;
      const clientId = import.meta.env.VITE_AZURE_CLIENT_ID as string | undefined;
      const tenantId = import.meta.env.VITE_AZURE_TENANT_ID as string | undefined;
      if (dataverseUrl && clientId && tenantId) {
        return { dataverseUrl, clientId, tenantId };
      }
      return null;
    }
    const parsed = JSON.parse(stored) as AppConfig;
    if (!parsed.dataverseUrl || !parsed.clientId || !parsed.tenantId) return null;
    return parsed;
  } catch {
    return null;
  }
}

export function saveConfig(config: AppConfig): void {
  localStorage.setItem(CONFIG_STORAGE_KEY, JSON.stringify(config));
}

export function clearConfig(): void {
  localStorage.removeItem(CONFIG_STORAGE_KEY);
}

export function createMsalConfig(config: AppConfig): Configuration {
  return {
    auth: {
      clientId: config.clientId,
      authority: `https://login.microsoftonline.com/${config.tenantId}`,
      redirectUri: window.location.origin,
    },
    cache: {
      cacheLocation: "sessionStorage",
      storeAuthStateInCookie: false,
    },
  };
}

export function getDataverseScopes(config: AppConfig): string[] {
  const baseUrl = config.dataverseUrl.replace(/\/$/, "");
  return [`${baseUrl}/.default`];
}
