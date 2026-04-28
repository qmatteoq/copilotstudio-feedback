import { useState } from "react";
import { AppConfig } from "../types";

interface Props {
  onSave: (config: AppConfig) => void;
}

const GUID_RE = /^[0-9a-f]{8}-([0-9a-f]{4}-){3}[0-9a-f]{12}$/i;
const DATAVERSE_URL_RE = /^https:\/\/.+\.dynamics\.com\/?$/i;

export default function ConfigPage({ onSave }: Props) {
  const [dataverseUrl, setDataverseUrl] = useState("");
  const [tenantId, setTenantId] = useState("");
  const [clientId, setClientId] = useState("");
  const [entitySetName, setEntitySetName] = useState("");
  const [showAdvanced, setShowAdvanced] = useState(false);
  const [errors, setErrors] = useState<Record<string, string>>({});

  const validate = (): boolean => {
    const newErrors: Record<string, string> = {};

    if (!dataverseUrl.trim()) {
      newErrors.dataverseUrl = "Environment URL is required.";
    } else if (!DATAVERSE_URL_RE.test(dataverseUrl.trim())) {
      newErrors.dataverseUrl =
        "Must be a valid Dataverse URL, e.g. https://org.crm.dynamics.com";
    }

    if (!tenantId.trim()) {
      newErrors.tenantId = "Tenant ID is required.";
    } else if (!GUID_RE.test(tenantId.trim())) {
      newErrors.tenantId = "Must be a valid GUID.";
    }

    if (!clientId.trim()) {
      newErrors.clientId = "Client ID is required.";
    } else if (!GUID_RE.test(clientId.trim())) {
      newErrors.clientId = "Must be a valid GUID.";
    }

    setErrors(newErrors);
    return Object.keys(newErrors).length === 0;
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (validate()) {
      onSave({
        dataverseUrl: dataverseUrl.trim().replace(/\/$/, ""),
        clientId: clientId.trim(),
        tenantId: tenantId.trim(),
        entitySetName: entitySetName.trim() || undefined,
      });
    }
  };

  return (
    <div className="config-container">
      <div className="config-card">
        <div className="config-header">
          <div className="config-logo">📊</div>
          <h1>Copilot Studio Feedback Viewer</h1>
          <p>Connect to your Power Platform environment to explore agent feedback.</p>
        </div>

        <form onSubmit={handleSubmit} className="config-form">
          <div className="form-group">
            <label htmlFor="dataverseUrl">Environment URL</label>
            <input
              id="dataverseUrl"
              type="url"
              placeholder="https://yourorg.crm.dynamics.com"
              value={dataverseUrl}
              onChange={(e) => setDataverseUrl(e.target.value)}
              className={errors.dataverseUrl ? "input-error" : ""}
              autoComplete="off"
              spellCheck={false}
            />
            {errors.dataverseUrl ? (
              <span className="field-error">{errors.dataverseUrl}</span>
            ) : (
              <small>Your Dataverse / Power Platform environment URL</small>
            )}
          </div>

          <div className="form-group">
            <label htmlFor="tenantId">Azure AD Tenant ID</label>
            <input
              id="tenantId"
              type="text"
              placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
              value={tenantId}
              onChange={(e) => setTenantId(e.target.value)}
              className={errors.tenantId ? "input-error" : ""}
              autoComplete="off"
              spellCheck={false}
            />
            {errors.tenantId && (
              <span className="field-error">{errors.tenantId}</span>
            )}
          </div>

          <div className="form-group">
            <label htmlFor="clientId">Azure AD Application (Client) ID</label>
            <input
              id="clientId"
              type="text"
              placeholder="xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx"
              value={clientId}
              onChange={(e) => setClientId(e.target.value)}
              className={errors.clientId ? "input-error" : ""}
              autoComplete="off"
              spellCheck={false}
            />
            {errors.clientId ? (
              <span className="field-error">{errors.clientId}</span>
            ) : (
              <small>
                Register an app in Azure AD with Dynamics CRM{" "}
                <code>user_impersonation</code> permission
              </small>
            )}
          </div>

          <div className="advanced-toggle">
            <button
              type="button"
              className="btn-link"
              onClick={() => setShowAdvanced((v) => !v)}
            >
              {showAdvanced ? "▲ Hide advanced options" : "▼ Advanced options"}
            </button>
          </div>

          {showAdvanced && (
            <div className="form-group">
              <label htmlFor="entitySetName">Entity Set Name override</label>
              <input
                id="entitySetName"
                type="text"
                placeholder="msdyn_conversationtranscripts (auto-discovered)"
                value={entitySetName}
                onChange={(e) => setEntitySetName(e.target.value)}
                autoComplete="off"
                spellCheck={false}
              />
              <small>
                Leave blank to auto-discover via Dataverse metadata. Set this only if
                auto-discovery fails — check the correct value in your environment's
                table metadata (logical name: <code>msdyn_conversationtranscript</code>).
              </small>
            </div>
          )}

          <button type="submit" className="btn-primary btn-full">
            Connect
          </button>
        </form>

        <div className="config-help">
          <h3>Setup guide</h3>
          <ol>
            <li>
              Go to <strong>Azure Portal → Azure Active Directory → App registrations</strong> and
              create a new registration.
            </li>
            <li>
              Under <strong>Authentication</strong>, add{" "}
              <code>{window.location.origin}</code> as a{" "}
              <strong>Single-page application</strong> redirect URI.
            </li>
            <li>
              Under <strong>API permissions</strong>, add{" "}
              <strong>Dynamics CRM → user_impersonation</strong> (delegated).
            </li>
            <li>
              Make sure the signed-in user has a <strong>Dataverse role</strong> with
              read access to the ConversationTranscript table.
            </li>
          </ol>
        </div>
      </div>
    </div>
  );
}
