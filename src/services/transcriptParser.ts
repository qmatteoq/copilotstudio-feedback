import { FeedbackItem, TranscriptActivity, TranscriptContent } from "../types";

/**
 * Parses a conversation transcript JSON string and extracts all feedback items.
 *
 * Feedback is identified by activities with:
 *   type === "invoke", name === "message/submitAction", value.actionName === "feedback"
 *
 * The agent message the feedback refers to is resolved via the activity's replyToId,
 * which points to a "message" activity from the bot (role 0).
 */
export function extractFeedback(
  transcriptId: string,
  content: string,
  agentName: string
): FeedbackItem[] {
  let parsed: TranscriptContent;
  try {
    parsed = JSON.parse(content) as TranscriptContent;
  } catch {
    return [];
  }

  if (!Array.isArray(parsed.activities)) return [];

  // Build lookup map: activity id → activity
  const activityMap = new Map<string, TranscriptActivity>();
  for (const activity of parsed.activities) {
    if (activity.id) {
      activityMap.set(activity.id, activity);
    }
  }

  const feedbackItems: FeedbackItem[] = [];
  let feedbackIndex = 0;

  for (const activity of parsed.activities) {
    if (
      activity.type === "invoke" &&
      activity.name === "message/submitAction" &&
      activity.value?.actionName === "feedback"
    ) {
      const feedbackText =
        activity.value.actionValue?.feedback?.feedbackText ?? "";
      const reaction = activity.value.actionValue?.reaction ?? "";

      // Resolve the agent message this feedback is attached to via replyToId
      let agentMessage = "";
      if (activity.replyToId) {
        const referenced = activityMap.get(activity.replyToId);
        if (referenced) {
          // Prefer the full spoken text; fall back to the display text
          agentMessage = referenced.speak ?? referenced.text ?? "";
        }
      }

      // Derive timestamp: timestampMs is in ms, timestamp is in seconds
      const timestampMs =
        activity.timestampMs ??
        (activity.timestamp != null ? activity.timestamp * 1000 : null);

      feedbackItems.push({
        id: `${transcriptId}-${activity.id ?? String(feedbackIndex)}`,
        agentName,
        feedbackText,
        reaction,
        agentMessage,
        timestamp: timestampMs != null ? new Date(timestampMs).toISOString() : "",
        transcriptId,
      });

      feedbackIndex++;
    }
  }

  return feedbackItems;
}
