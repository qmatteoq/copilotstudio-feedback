import { Fragment, useState } from "react";
import { FeedbackItem } from "../types";

interface Props {
  items: FeedbackItem[];
}

function ReactionBadge({ reaction }: { reaction: string }) {
  if (reaction === "like") {
    return <span className="badge badge-like">👍 Like</span>;
  }
  if (reaction === "dislike") {
    return <span className="badge badge-dislike">👎 Dislike</span>;
  }
  return <span className="badge badge-neutral">{reaction || "—"}</span>;
}

function truncate(text: string, max: number): string {
  if (text.length <= max) return text;
  return text.slice(0, max) + "…";
}

function formatDate(iso: string): string {
  if (!iso) return "—";
  try {
    return new Date(iso).toLocaleString(undefined, {
      dateStyle: "short",
      timeStyle: "short",
    });
  } catch {
    return iso;
  }
}

export default function FeedbackTable({ items }: Props) {
  const [expandedId, setExpandedId] = useState<string | null>(null);

  if (items.length === 0) {
    return (
      <div className="empty-state">
        <span className="empty-icon">🔍</span>
        <p>No feedback found matching your filters.</p>
      </div>
    );
  }

  return (
    <div className="table-wrapper">
      <table className="feedback-table">
        <thead>
          <tr>
            <th style={{ width: 110 }}>Reaction</th>
            <th style={{ width: 160 }}>Agent</th>
            <th>Feedback</th>
            <th>Agent Message</th>
            <th style={{ width: 130 }}>Timestamp</th>
          </tr>
        </thead>
        <tbody>
          {items.map((item) => {
            const isExpanded = expandedId === item.id;
            return (
              <Fragment key={item.id}>
                <tr
                  className={`feedback-row${isExpanded ? " row-expanded" : ""}`}
                  onClick={() => setExpandedId(isExpanded ? null : item.id)}
                  title="Click to expand"
                  aria-expanded={isExpanded}
                >
                  <td>
                    <ReactionBadge reaction={item.reaction} />
                  </td>
                  <td className="cell-agent">{item.agentName}</td>
                  <td className="cell-text">
                    {truncate(item.feedbackText, 140) || (
                      <em className="muted">No text provided</em>
                    )}
                  </td>
                  <td className="cell-text">
                    {truncate(item.agentMessage, 140) || (
                      <em className="muted">Not available</em>
                    )}
                  </td>
                  <td className="cell-timestamp">{formatDate(item.timestamp)}</td>
                </tr>

                {isExpanded && (
                  <tr className="detail-row">
                    <td colSpan={5}>
                      <div className="detail-content">
                        <div className="detail-section">
                          <h4>Feedback</h4>
                          <p>{item.feedbackText || "No text provided."}</p>
                        </div>
                        <div className="detail-section">
                          <h4>Agent Message</h4>
                          <p>{item.agentMessage || "Not available."}</p>
                        </div>
                        <div className="detail-meta">
                          <span>
                            <strong>Agent:</strong> {item.agentName}
                          </span>
                          <span>
                            <strong>Reaction:</strong>{" "}
                            <ReactionBadge reaction={item.reaction} />
                          </span>
                          <span>
                            <strong>Time:</strong> {formatDate(item.timestamp)}
                          </span>
                          <span className="muted">
                            Transcript: {item.transcriptId}
                          </span>
                        </div>
                      </div>
                    </td>
                  </tr>
                )}
              </Fragment>
            );
          })}
        </tbody>
      </table>
    </div>
  );
}
