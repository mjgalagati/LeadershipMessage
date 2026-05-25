import * as React from 'react';
import { PrimaryButton, Icon } from '@fluentui/react';
import { IAlert } from '../models/IAlert';
import styles from './AlertDetailsPanel.module.scss';

interface AlertDetailsPanelProps {
  alert?: IAlert;
  isOpen: boolean;
  onDismiss: () => void;
  onEdit: (alert: IAlert) => void;
  currentUserId: number;
}

const AlertDetailsPanel: React.FC<AlertDetailsPanelProps> = ({
  alert, isOpen, onDismiss, onEdit, currentUserId,
}) => {
  if (!isOpen || !alert) return null; // eslint-disable-line @rushstack/no-new-null

  const isExpired = alert.Status === 'Expired';
  const isAuthor = alert.Author?.Id === currentUserId;
  const canEdit = !isExpired && isAuthor;
  const hasBanner = !!alert.BannerImageUrl;

  const closeBtnClass = [
    styles.closeBtn,
    hasBanner ? styles.closeBtnOnBanner : styles.closeBtnNoImage,
  ].join(" ");

  const getFileIcon = (fileName: string): string => {
    const ext = fileName.split('.').pop()?.toLowerCase() ?? '';
    if (ext === 'pdf') return 'PDF';
    if (ext === 'doc' || ext === 'docx') return 'WordDocument';
    if (ext === 'xls' || ext === 'xlsx' || ext === 'csv') return 'ExcelDocument';
    if (ext === 'ppt' || ext === 'pptx') return 'PowerPointDocument';
    if (ext === 'jpg' || ext === 'jpeg' || ext === 'png' || ext === 'gif' || ext === 'webp') return 'FileImage';
    if (ext === 'mp4') return 'Video';
    if (ext === 'mp3') return 'Music';
    return 'Attach';
  };

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const priorityBadgeClass = (): string => {
    if (alert.Priority === 'Critical') return styles.priorityCritical;
    if (alert.Priority === 'High') return styles.priorityHigh;
    if (alert.Priority === 'Medium') return styles.priorityMedium;
    return styles.priorityLow;
  };

  const statusBadgeClass = (): string => {
    if (alert.Status === 'Active') return styles.statusActive;
    if (alert.Status === 'Scheduled') return styles.statusScheduled;
    if (alert.Status === 'Expired') return styles.statusExpired;
    return styles.statusDraft;
  };

  return (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Close button — always rendered, anchored to modal top-right ── */}
        <button className={closeBtnClass} onClick={onDismiss} aria-label="Close">
          <Icon iconName="Cancel" />
        </button>

        {/* ── Hero Banner ── */}
        {hasBanner && (
          <div className={styles.bannerContainer}>
            <img
              src={alert.BannerImageUrl}
              alt={alert.Title ?? 'Alert Banner'}
              className={styles.bannerImage}
            />
            <div className={styles.bannerOverlayContent}>
              <div className={styles.bannerTypeBadgeRow}>
                <span className={styles.bannerTypeBadge}>{alert.AlertType ?? 'General'}</span>
                <span className={styles.bannerTypeBadge}>{alert.Priority}</span>
              </div>
              <h1 className={styles.bannerTitle}>{alert.Title ?? 'Untitled'}</h1>
              {alert.Author && (
                <span className={styles.bannerAuthor}>
                  <Icon iconName="Contact" />
                  {alert.Author.Title}
                </span>
              )}
              <span className={styles.bannerPublishDate}>
                {alert.PublishDate.toLocaleDateString('en-US', {
                  weekday: 'short', year: 'numeric', month: 'short', day: 'numeric',
                })}
              </span>
            </div>
            {isExpired && (
              <div className={styles.expiredOverlay}>
                <span>Expired</span>
              </div>
            )}
          </div>
        )}

        {/* ── Scrollable article body ── */}
        <div className={styles.scrollBody}>
          <div className={styles.content}>

            {/* When no banner: show title/meta in content area */}
            {!hasBanner && (
              <div className={styles.headerNoBanner}>
                <div className={styles.titleSection}>
                  <div className={styles.typeBadgeRow}>
                    <span className={styles.typeBadge}>{alert.AlertType ?? 'General'}</span>
                    <span className={[styles.priorityBadge, priorityBadgeClass()].join(" ")}>
                      {alert.Priority}
                    </span>
                  </div>
                  <h1 className={styles.alertTitle}>{alert.Title ?? 'Untitled'}</h1>
                  {alert.Author && (
                    <span className={styles.authorLabel}>
                      <Icon iconName="Contact" />
                      {alert.Author.Title}
                    </span>
                  )}
                </div>
              </div>
            )}

            {/* Edit / locked controls */}
            {canEdit && (
              <div className={styles.articleActions}>
                <PrimaryButton
                  text="Edit"
                  iconProps={{ iconName: 'Edit' }}
                  onClick={() => onEdit(alert)}
                  className={styles.editButton}
                />
              </div>
            )}
            {!canEdit && isExpired && (
              <div className={styles.articleActions}>
                <div className={styles.lockedNote}>
                  <Icon iconName="Lock" />
                  <span>Edit via SharePoint list</span>
                </div>
              </div>
            )}
            {!canEdit && !isExpired && !isAuthor && (
              <div className={styles.articleActions}>
                <div className={styles.lockedNote}>
                  <Icon iconName="Lock" />
                  <span>Only the creator can edit</span>
                </div>
              </div>
            )}

            {/* Body */}
            {alert.Body && (
              <div
                className={styles.bodyBox}
                dangerouslySetInnerHTML={{ __html: alert.Body }}
              />
            )}

            {/* Attachments */}
            {alert.Attachments && alert.Attachments.length > 0 && (
              <div className={styles.infoCard}>
                <div className={styles.cardHeader}>
                  <Icon iconName="Attach" className={styles.cardIcon} />
                  <h3>Attachments ({alert.Attachments.length})</h3>
                </div>
                <div className={styles.cardContent}>
                  <div className={styles.attachmentList}>
                    {alert.Attachments.map((att, idx) => (
                      <a
                        key={idx}
                        href={att.ServerRelativeUrl}
                        target="_blank"
                        rel="noreferrer"
                        className={styles.attachmentCard}
                      >
                        <div className={styles.attachmentIconWrapper}>
                          <Icon iconName={getFileIcon(att.FileName)} className={styles.attachmentTypeIcon} />
                        </div>
                        <div className={styles.attachmentDetails}>
                          <span className={styles.attachmentName}>{att.FileName}</span>
                          <span className={styles.attachmentAction}>Click to open</span>
                        </div>
                        <Icon iconName="OpenInNewWindow" className={styles.attachmentOpenIcon} />
                      </a>
                    ))}
                  </div>
                </div>
              </div>
            )}

            {/* Status footer */}
            <div className={styles.statusContainer}>
              <span className={[styles.statusBadge, statusBadgeClass()].join(" ")}>
                {alert.Status ?? 'Draft'}
              </span>
            </div>

          </div>
        </div>
      </div>
    </div>
  );
};

export default AlertDetailsPanel;