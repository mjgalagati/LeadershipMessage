import * as React from "react";
import * as ReactDOM from "react-dom";
import { Icon } from "@fluentui/react";
import { ILeadershipMessage } from "../models/ILeadershipMessage";
import { IExecutive } from "../models/IExecutive";
import styles from "./ViewAllModal.module.scss";

interface ViewAllModalProps {
  messages: ILeadershipMessage[];
  executives: IExecutive[];
  isOpen: boolean;
  onDismiss: () => void;
  onSelectMessage: (message: ILeadershipMessage) => void;
  onAddMessage: () => void;
  onAddExecutive: () => void;
  onEditExecutive: (exec: IExecutive) => void;
  isEditor: boolean;
}

const stripHtml = (html: string): string => {
  const tmp = document.createElement("div");
  tmp.innerHTML = html;
  const text = tmp.textContent || tmp.innerText || "";
  return text.replace(/\u00a0/g, " ").trim();
};

const truncate = (text: string, max: number): string =>
  text.length <= max ? text : text.slice(0, max).replace(/\s+$/, "") + "…";

const ViewAllModal: React.FC<ViewAllModalProps> = ({
  messages, executives, isOpen, onDismiss,
  onSelectMessage, onAddMessage, onAddExecutive, onEditExecutive, isEditor,
}) => {
  const [activeTab, setActiveTab] = React.useState<"messages" | "executives">("messages");

  if (!isOpen) return <></>;

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  return ReactDOM.createPortal(
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        <div className={styles.modalHeader}>
          <h2 className={styles.modalTitle}>All Content</h2>
          <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
            <Icon iconName="Cancel" />
          </button>
        </div>

        {/* ── Tabs ── */}
        <div className={styles.tabs}>
          <button
            className={activeTab === "messages" ? styles.tabActive : styles.tab}
            onClick={() => setActiveTab("messages")}
          >
            Messages ({messages.length})
          </button>
          <button
            className={activeTab === "executives" ? styles.tabActive : styles.tab}
            onClick={() => setActiveTab("executives")}
          >
            Executives ({executives.length})
          </button>
        </div>

        {/* ── Action bar ── */}
        {isEditor && (
          <div className={styles.actionBar}>
            {activeTab === "messages" ? (
              <button className={styles.addBtn} onClick={onAddMessage}>
                <Icon iconName="Add" /> Add Message
              </button>
            ) : (
              <button className={styles.addBtn} onClick={onAddExecutive}>
                <Icon iconName="Add" /> Add Executive
              </button>
            )}
          </div>
        )}

        {/* ── Content ── */}
        <div className={styles.scrollBody}>
          {activeTab === "messages" && (
            messages.length === 0
              ? <p className={styles.emptyMsg}>No messages yet.</p>
              : (
                <ul className={styles.list}>
                  {messages.map(msg => {
                    const exec = msg.Executive;
                    const preview = truncate(stripHtml(msg.Body ?? ""), 100);
                    return (
                      <li key={msg.Id} className={styles.listRow} onClick={() => onSelectMessage(msg)}>
                        <div className={styles.rowAvatar}>
                          {exec?.PhotoUrl ? (
                            <img
                              src={exec.PhotoUrl}
                              alt={exec.Title}
                              className={styles.rowAvatarImg}
                              style={{ objectPosition: `${exec.PhotoFocalX}% ${exec.PhotoFocalY}%` }}
                            />
                          ) : (
                            <div className={styles.rowAvatarFallback}>
                              {exec ? exec.Title.charAt(0).toUpperCase() : "?"}
                            </div>
                          )}
                        </div>
                        <div className={styles.rowContent}>
                          <span className={styles.rowTitle}>{msg.Title}</span>
                          {exec && (
                            <span className={styles.rowMeta}>{exec.Title} · {exec.Position}</span>
                          )}
                          {preview && <span className={styles.rowPreview}>{preview}</span>}
                        </div>
                        <div className={styles.rowDate}>
                          {msg.Created?.toLocaleDateString("en-US", { month: "short", day: "numeric", year: "numeric" })}
                        </div>
                        <Icon iconName="ChevronRight" className={styles.rowChevron} />
                      </li>
                    );
                  })}
                </ul>
              )
          )}

          {activeTab === "executives" && (
            executives.length === 0
              ? <p className={styles.emptyMsg}>No executives yet.</p>
              : (
                <ul className={styles.list}>
                  {executives.map(exec => (
                    <li
                      key={exec.Id}
                      className={styles.listRow}
                      onClick={isEditor ? () => onEditExecutive(exec) : undefined}
                      style={isEditor ? undefined : { cursor: 'default' }}
                    >
                      <div className={styles.rowAvatar}>
                        {exec.PhotoUrl ? (
                          <img
                            src={exec.PhotoUrl}
                            alt={exec.Title}
                            className={styles.rowAvatarImg}
                            style={{ objectPosition: `${exec.PhotoFocalX}% ${exec.PhotoFocalY}%` }}
                          />
                        ) : (
                          <div className={styles.rowAvatarFallback}>
                            {exec.Title.charAt(0).toUpperCase()}
                          </div>
                        )}
                      </div>
                      <div className={styles.rowContent}>
                        <span className={styles.rowTitle}>{exec.Title}</span>
                        <span className={styles.rowMeta}>{exec.Position}</span>
                      </div>
                      {isEditor && <Icon iconName="Edit" className={styles.rowChevron} />}
                    </li>
                  ))}
                </ul>
              )
          )}
        </div>

      </div>
    </div>,
    document.body
  );
};

export default ViewAllModal;
