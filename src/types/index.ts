export interface AppConfig {
  dataverseUrl: string;
  clientId: string;
  tenantId: string;
  /** Friendly display name for the environment (optional). */
  environmentName?: string;
  /** Optional override for the OData entity set name (auto-discovered if omitted). */
  entitySetName?: string;
}

export interface FeedbackItem {
  id: string;
  agentName: string;
  feedbackText: string;
  reaction: string;
  agentMessage: string;
  timestamp: string;
  transcriptId: string;
}

// Field names differ between environments depending on which solution installed the table.
// We use a generic record and resolve fields at runtime via TranscriptSchema.
export type ConversationTranscript = Record<string, unknown>;

export interface TranscriptSchema {
  entitySetName: string;
  logicalName: string;
  idField: string;
  nameField: string;
  contentField: string;
  botAnnotationPrefixes: string[];
}

export interface TranscriptActivity {
  id?: string;
  type: string;
  name?: string;
  timestamp?: number;
  timestampMs?: number;
  from?: {
    id: string;
    role: number;
  };
  text?: string;
  speak?: string;
  replyToId?: string;
  value?: {
    actionName?: string;
    actionValue?: {
      feedback?: {
        feedbackText?: string;
      };
      reaction?: string;
    };
    [key: string]: unknown;
  };
}

export interface TranscriptContent {
  activities: TranscriptActivity[];
}

export interface ODataResponse<T> {
  value: T[];
  "@odata.nextLink"?: string;
}
