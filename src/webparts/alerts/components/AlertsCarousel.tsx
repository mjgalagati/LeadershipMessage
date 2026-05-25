import * as React from "react";
import { useEffect, useState } from "react";
import styles from "./AlertsCarousel.module.scss";
import { IAlertsCarouselProps } from "./IAlertsCarouselProps";
import { AlertsService } from "../services/AlertsService";
import { IAlert } from "../models/IAlert";
import AlertDetailsPanel from "./AlertDetailsPanel";
import ViewAllPanel from "./ViewAllPanel";
import AddEditAlert from "./AddEditAlert";

const stripHtml = (html: string): string =>
  html.replace(/<[^>]+>/g, "").trim();

const getPriorityStyle = (priority: string): string => {
  switch (priority) {
    case "Critical": return styles.priorityCritical;
    case "High":     return styles.priorityHigh;
    case "Medium":   return styles.priorityMedium;
    default:         return styles.priorityLow;
  }
};

const getPriorityLabel = (priority: string): string => {
  switch (priority) {
    case "Critical": return "CRITICAL!";
    case "High":     return "HIGH PRIORITY";
    case "Medium":   return "MEDIUM";
    default:         return "LOW";
  }
};

const AlertsCarousel = (props: IAlertsCarouselProps): JSX.Element => {
  const [alerts, setAlerts] = useState<IAlert[]>([]);
  const [selectedAlert, setSelectedAlert] = useState<IAlert | undefined>(undefined);
  const [isDetailsOpen, setIsDetailsOpen] = useState(false);
  const [isViewAllOpen, setIsViewAllOpen] = useState(false);
  const [isAddEditOpen, setIsAddEditOpen] = useState(false);
  const [addEditMode, setAddEditMode] = useState<"add" | "edit">("add");
  const [alertToEdit, setAlertToEdit] = useState<IAlert | undefined>(undefined);
  const [currentIndex, setCurrentIndex] = useState(0);

  const loadAlerts = async (): Promise<void> => {
    try {
      const service = new AlertsService(props.context, props.sourceList);
      const data = await service.getAlerts(props.currentUserLogin);
      setAlerts(data);
    } catch (err) {
      console.error("Failed to load alerts", err);
    }
  };

  useEffect(() => {
    if (!props.sourceList) return;
    loadAlerts().catch(console.error);
  }, [props.sourceList, props.context, props.currentUserLogin]);

  const activeAlerts = alerts.filter(a => a.Status?.trim() === "Active");
  const current = activeAlerts[currentIndex];
  const isMultiple = activeAlerts.length > 1;

  // Auto-rotate if multiple active alerts
  useEffect(() => {
    if (!isMultiple) return;
    const timer = setInterval(() => {
      setCurrentIndex(prev => (prev + 1) % activeAlerts.length);
    }, 8000);
    return () => clearInterval(timer);
  }, [isMultiple, activeAlerts.length]);

  const openDetailsFromCarousel = (alert: IAlert): void => {
    setSelectedAlert(alert);
    setIsDetailsOpen(true);
  };

  const openDetailsFromViewAll = (alert: IAlert): void => {
    setSelectedAlert(alert);
    setIsDetailsOpen(true);
  };

  const closeDetails = (): void => {
    setIsDetailsOpen(false);
    setSelectedAlert(undefined);
  };

  const openAddAlert = (): void => {
    setAddEditMode("add");
    setAlertToEdit(undefined);
    setIsAddEditOpen(true);
  };

  const openEditAlert = (alert: IAlert): void => {
    setIsDetailsOpen(false);
    setAddEditMode("edit");
    setAlertToEdit(alert);
    setIsAddEditOpen(true);
  };

  const closeAddEditPanel = (): void => {
    setIsAddEditOpen(false);
    if (alertToEdit) {
      setSelectedAlert(alertToEdit);
      setIsDetailsOpen(true);
    }
    setAlertToEdit(undefined);
  };

  const handleSaveAlert = async (
    alertData: Partial<IAlert>,
    bannerFile?: File,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> => {
    const service = new AlertsService(props.context, props.sourceList);
    try {
      if (addEditMode === "add") {
        await service.addAlert(alertData, bannerFile, attachments);
      } else if (addEditMode === "edit" && alertToEdit) {
        await service.updateAlert(alertToEdit.Id, alertData, bannerFile, attachments, deletedAttachmentNames);
      }
      setIsAddEditOpen(false);
      setAlertToEdit(undefined);
      setSelectedAlert(undefined);
      setIsDetailsOpen(false);
      setIsViewAllOpen(false);
      await loadAlerts();
    } catch (err) {
      console.error("Failed to save alert:", err);
      throw err;
    }
  };

  const renderPanels = (): JSX.Element => (
    <>
      <ViewAllPanel
        alerts={alerts}
        isOpen={isViewAllOpen}
        onDismiss={() => setIsViewAllOpen(false)}
        onSelectAlert={openDetailsFromViewAll}
        onAddAlert={openAddAlert}
      />
      <AlertDetailsPanel
        alert={selectedAlert}
        isOpen={isDetailsOpen}
        onDismiss={closeDetails}
        onEdit={openEditAlert}
        currentUserId={props.currentUserId}
      />
      <AddEditAlert
        isOpen={isAddEditOpen}
        mode={addEditMode}
        alert={alertToEdit}
        context={props.context}
        onDismiss={closeAddEditPanel}
        onSave={handleSaveAlert}
      />
    </>
  );

  // ── No active alerts fallback ──
  if (!activeAlerts.length) {
    return (
      <div className={styles.wrapper}>
        <div className={styles.emptyState}>
          <div className={styles.emptyHeader}>
            <span className={styles.sectionLabel}>ALERTS</span>
            <button className={styles.viewAllLink} onClick={() => setIsViewAllOpen(true)}>
              View all
            </button>
          </div>
          <div className={styles.emptyBody}>
            <span className={styles.emptyIcon}>📢</span>
            {!props.sourceList ? (
              <p>Please select a list in the web part settings.</p>
            ) : (
              <p>No active alerts</p>
            )}
            {props.sourceList && (
              <button className={styles.addBtn} onClick={openAddAlert}>
                + Add Alert
              </button>
            )}
          </div>
        </div>
        {props.sourceList && renderPanels()}
      </div>
    );
  }

  const bodyText = stripHtml(current.Body ?? "");

  return (
    <div className={styles.wrapper}>
      <div className={styles.card}>

        {/* ── Header: ALERTS + View all ── */}
        <div className={styles.cardHeader}>
          <span className={styles.sectionLabel}>ALERTS</span>
          <button className={styles.viewAllLink} onClick={() => setIsViewAllOpen(true)}>
            View all
          </button>
        </div>

        {/* ── Priority badge + date ── */}
        <div className={styles.metaRow}>
          <span className={styles.priorityBadge + " " + getPriorityStyle(current.Priority)}>
            {current.Priority === "Critical" && (
              <span className={styles.pulseDot} />
            )}
            {getPriorityLabel(current.Priority)}
          </span>
          <span className={styles.metaDate}>
            {current.PublishDate.toLocaleDateString("en-US", {
              month: "short", day: "numeric", year: "numeric",
            })}
          </span>
        </div>

        {/* ── Title ── */}
        <h3 className={styles.title}>{current.Title}</h3>

        {/* ── Body text ── */}
        {bodyText && (
          <p className={styles.body}>{bodyText}</p>
        )}

        {/* ── Footer: dots + actions ── */}
        <div className={styles.cardFooter}>
          {isMultiple && (
            <div className={styles.dots}>
              {activeAlerts.map((_, i) => (
                <button
                  key={i}
                  className={styles.dot + (i === currentIndex ? " " + styles.dotActive : "")}
                  onClick={() => setCurrentIndex(i)}
                />
              ))}
            </div>
          )}
          <div className={styles.footerActions}>
            <button
              className={styles.readMoreBtn}
              onClick={() => openDetailsFromCarousel(current)}
            >
              Read More →
            </button>
            <button
              className={styles.addBtnFooter}
              onClick={openAddAlert}
            >
              + Add
            </button>
          </div>
        </div>

      </div>

      {renderPanels()}
    </div>
  );
};

export default AlertsCarousel;