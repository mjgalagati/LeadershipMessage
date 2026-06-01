import * as React from "react";
import { Icon } from "@fluentui/react";
import { ILeadershipMessage } from "../models/ILeadershipMessage";
import styles from "./MessageDetailModal.module.scss";

interface MessageDetailModalProps {
  message?: ILeadershipMessage;
  isOpen: boolean;
  onDismiss: () => void;
  onEdit: (message: ILeadershipMessage) => void;
  currentUserId: number;
}

const getFileIcon = (fileName: string): string => {
  const ext = fileName.split(".").pop()?.toLowerCase() ?? "";
  if (ext === "pdf") return "PDF";
  if (ext === "doc" || ext === "docx") return "WordDocument";
  if (ext === "xls" || ext === "xlsx" || ext === "csv") return "ExcelDocument";
  if (ext === "ppt" || ext === "pptx") return "PowerPointDocument";
  if (["jpg", "jpeg", "png", "gif", "webp"].includes(ext)) return "FileImage";
  if (ext === "mp4") return "Video";
  if (ext === "mp3") return "Music";
  return "Attach";
};

const MessageDetailModal: React.FC<MessageDetailModalProps> = ({
  message, isOpen, onDismiss, onEdit, currentUserId,
}) => {
  if (!isOpen || !message) return <></>;

  const exec = message.Executive;
  const isAuthor = message.Author?.Id === currentUserId;

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  return (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
          <Icon iconName="Cancel" />
        </button>

        <div className={styles.scrollBody}>
          <div className={styles.content}>

            {/* ── Executive header ── */}
            <div className={styles.execHeader}>
              {exec?.PhotoUrl ? (
                <img
                  src={exec.PhotoUrl}
                  alt={exec.Title}
                  className={styles.avatar}
                  style={{ objectPosition: `${exec.PhotoFocalX}% ${exec.PhotoFocalY}%` }}
                />
              ) : (
                <div className={styles.avatarFallback}>
                  {exec ? exec.Title.charAt(0).toUpperCase() : "?"}
                </div>
              )}

              <div className={styles.execInfo}>
                {exec && (
                  <>
                    <span className={styles.execName}>{exec.Title}</span>
                    <span className={styles.execPosition}>{exec.Position}</span>
                  </>
                )}
                {message.Created && (
                  <span className={styles.publishDate}>
                    {message.Created.toLocaleDateString("en-US", {
                      weekday: "short", year: "numeric", month: "short", day: "numeric",
                    })}
                  </span>
                )}
              </div>

              {isAuthor && (
                <button
                  className={styles.editBtn}
                  onClick={() => onEdit(message)}
                  aria-label="Edit message"
                >
                  <Icon iconName="Edit" />
                  <span>Edit</span>
                </button>
              )}
            </div>

            {/* ── Title ── */}
            <h1 className={styles.title}>{message.Title}</h1>

            {/* ── Body ── */}
            {message.Body && (
              <div
                className={styles.bodyBox}
                dangerouslySetInnerHTML={{ __html: message.Body }}
              />
            )}

            {/* ── Attachments ── */}
            {message.Attachments && message.Attachments.length > 0 && (
              <div className={styles.attachmentsSection}>
                <h3 className={styles.attachmentsHeading}>
                  <Icon iconName="Attach" />
                  Attachments ({message.Attachments.length})
                </h3>
                <div className={styles.attachmentList}>
                  {message.Attachments.map((att, idx) => (
                    <a
                      key={idx}
                      href={att.ServerRelativeUrl}
                      target="_blank"
                      rel="noreferrer"
                      className={styles.attachmentCard}
                    >
                      <div className={styles.attachmentIconWrap}>
                        <Icon iconName={getFileIcon(att.FileName)} className={styles.attachmentTypeIcon} />
                      </div>
                      <div className={styles.attachmentDetails}>
                        <span className={styles.attachmentName}>{att.FileName}</span>
                        <span className={styles.attachmentHint}>Click to open</span>
                      </div>
                      <Icon iconName="OpenInNewWindow" className={styles.attachmentOpenIcon} />
                    </a>
                  ))}
                </div>
              </div>
            )}

          </div>
        </div>

      </div>
    </div>
  );
};

export default MessageDetailModal;
