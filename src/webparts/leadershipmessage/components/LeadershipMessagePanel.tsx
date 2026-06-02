import * as React from "react";
import { useState, useEffect } from "react";
import styles from "./LeadershipMessagePanel.module.scss";
import { ILeadershipMessagePanelProps } from "./ILeadershipMessagePanelProps";
import { LeadershipMessageService } from "../services/LeadershipMessageService";
import { ExecutiveService } from "../services/ExecutiveService";
import { ILeadershipMessage } from "../models/ILeadershipMessage";
import { IExecutive } from "../models/IExecutive";
import MessageDetailModal from "./MessageDetailModal";
import ViewAllModal from "./ViewAllModal";
import AddEditMessageModal from "./AddEditMessageModal";
import AddEditExecutiveModal from "./AddEditExecutiveModal";

const stripHtml = (html: string): string => html.replace(/<[^>]+>/g, "").trim();

const truncate = (text: string, max: number): string =>
  text.length <= max ? text : text.slice(0, max).replace(/\s+$/, "") + "…";

const LeadershipMessagePanel = (props: ILeadershipMessagePanelProps): JSX.Element => {
  const [messages, setMessages] = useState<ILeadershipMessage[]>([]);
  const [executives, setExecutives] = useState<IExecutive[]>([]);
  const [detailMessage, setDetailMessage] = useState<ILeadershipMessage | undefined>(undefined);
  const [isDetailOpen, setIsDetailOpen] = useState(false);
  const [isViewAllOpen, setIsViewAllOpen] = useState(false);
  const [isAddEditMessageOpen, setIsAddEditMessageOpen] = useState(false);
  const [isAddEditExecutiveOpen, setIsAddEditExecutiveOpen] = useState(false);
  const [messageMode, setMessageMode] = useState<"add" | "edit">("add");
  const [messageToEdit, setMessageToEdit] = useState<ILeadershipMessage | undefined>(undefined);
  const [executiveToEdit, setExecutiveToEdit] = useState<IExecutive | undefined>(undefined);

  const loadData = async (): Promise<void> => {
    let loadedMessages: ILeadershipMessage[] = [];
    let loadedExecutives: IExecutive[] = [];

    await Promise.all([
      props.sourceList
        ? new LeadershipMessageService(props.context, props.sourceList)
            .getMessages()
            .then(m => { loadedMessages = m; })
            .catch(err => console.error("Failed to load messages:", err))
        : Promise.resolve(),
      props.executiveList
        ? new ExecutiveService(props.context, props.executiveList)
            .getExecutives()
            .then(e => { loadedExecutives = e; })
            .catch(err => console.error("Failed to load executives:", err))
        : Promise.resolve(),
    ]);

    // PhotoUrl is a Hyperlink column that cannot be expanded through a lookup.
    // Merge full executive data (photo, focal points) from the executives fetch.
    const enriched = loadedMessages.map(msg => {
      if (!msg.Executive?.Id) return msg;
      const full = loadedExecutives.find(e => e.Id === msg.Executive!.Id);
      return full ? { ...msg, Executive: full } : msg;
    });

    setMessages(enriched);
    setExecutives(loadedExecutives);
  };

  useEffect(() => {
    loadData().catch(console.error);
  }, [props.sourceList, props.executiveList, props.context]);

  const latest = messages[0];
  const label = props.sectionTitle || "EXECUTIVE MESSAGE";

  const openDetail = (msg: ILeadershipMessage): void => {
    setDetailMessage(msg);
    setIsDetailOpen(true);
  };

  const closeDetail = (): void => {
    setIsDetailOpen(false);
    setDetailMessage(undefined);
  };

  const openAddMessage = (): void => {
    setIsViewAllOpen(false);
    setMessageToEdit(undefined);
    setMessageMode("add");
    setIsAddEditMessageOpen(true);
  };

  const openEditMessage = (msg: ILeadershipMessage): void => {
    setIsDetailOpen(false);
    setDetailMessage(undefined);
    setMessageToEdit(msg);
    setMessageMode("edit");
    setIsAddEditMessageOpen(true);
  };

  const openAddExecutive = (): void => {
    setIsViewAllOpen(false);
    setExecutiveToEdit(undefined);
    setIsAddEditExecutiveOpen(true);
  };

  const openEditExecutive = (exec: IExecutive): void => {
    setIsViewAllOpen(false);
    setExecutiveToEdit(exec);
    setIsAddEditExecutiveOpen(true);
  };

  const handleSaveMessage = async (
    data: Partial<ILeadershipMessage>,
    attachments?: File[],
    deletedAttachmentNames?: string[]
  ): Promise<void> => {
    const svc = new LeadershipMessageService(props.context, props.sourceList);
    if (messageMode === "add") {
      await svc.addMessage(data, attachments);
    } else if (messageToEdit) {
      await svc.updateMessage(messageToEdit.Id, data, attachments, deletedAttachmentNames);
    }
    setIsAddEditMessageOpen(false);
    setMessageToEdit(undefined);
    await loadData();
  };

  const handleSaveExecutive = async (data: Partial<IExecutive>, photoFile?: File): Promise<void> => {
    const svc = new ExecutiveService(props.context, props.executiveList);
    if (executiveToEdit) {
      await svc.updateExecutive(executiveToEdit.Id, data, photoFile);
    } else {
      await svc.addExecutive(data, photoFile);
    }
    setIsAddEditExecutiveOpen(false);
    setExecutiveToEdit(undefined);
    await loadData();
  };

  const renderModals = (): JSX.Element => (
    <>
      <MessageDetailModal
        message={detailMessage}
        isOpen={isDetailOpen}
        onDismiss={closeDetail}
        onEdit={openEditMessage}
        isEditor={props.isEditor}
      />
      <ViewAllModal
        messages={messages}
        executives={executives}
        isOpen={isViewAllOpen}
        onDismiss={() => setIsViewAllOpen(false)}
        onSelectMessage={openDetail}
        onAddMessage={openAddMessage}
        onAddExecutive={openAddExecutive}
        onEditExecutive={openEditExecutive}
        isEditor={props.isEditor}
      />
      <AddEditMessageModal
        isOpen={isAddEditMessageOpen}
        mode={messageMode}
        message={messageToEdit}
        executives={executives}
        context={props.context}
        onDismiss={() => { setIsAddEditMessageOpen(false); setMessageToEdit(undefined); }}
        onSave={handleSaveMessage}
      />
      <AddEditExecutiveModal
        isOpen={isAddEditExecutiveOpen}
        executive={executiveToEdit}
        onDismiss={() => { setIsAddEditExecutiveOpen(false); setExecutiveToEdit(undefined); }}
        onSave={handleSaveExecutive}
      />
    </>
  );

  if (!latest) {
    return (
      <div className={styles.wrapper}>
        <div className={styles.emptyState}>
          <div className={styles.emptyHeader}>
            <span className={styles.sectionLabel}>{label}</span>
            <button className={styles.viewAllLink} onClick={() => setIsViewAllOpen(true)}>View all</button>
          </div>
          <div className={styles.emptyBody}>
            {!props.sourceList
              ? <p>Configure source lists in web part settings.</p>
              : <p>No messages yet.</p>
            }
            {props.sourceList && props.isEditor && (
              <button className={styles.addBtn} onClick={openAddMessage}>+ Add Message</button>
            )}
          </div>
        </div>
        {props.sourceList && renderModals()}
      </div>
    );
  }

  const exec = latest.Executive;
  const pullQuote = truncate(stripHtml(latest.Body ?? ""), 320);

  return (
    <div className={styles.wrapper}>
      <div className={styles.card}>

        <div className={styles.cardHeader}>
          <span className={styles.sectionLabel}>{label}</span>
          <button className={styles.viewAllLink} onClick={() => setIsViewAllOpen(true)}>View all</button>
        </div>

        <div className={styles.messageBody}>
          <div className={styles.avatarCol}>
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
          </div>

          <div className={styles.contentCol}>
            {exec && (
              <div className={styles.execMeta}>
                <span className={styles.execName}>{exec.Title}</span>
                <span className={styles.execPosition}>{exec.Position}</span>
              </div>
            )}
            <h3 className={styles.messageTitle}>{latest.Title}</h3>
            {pullQuote && <p className={styles.pullQuote}>{pullQuote}</p>}
          </div>
        </div>

        <div className={styles.cardFooter}>
          <button className={styles.readMoreBtn} onClick={() => openDetail(latest)}>Read More →</button>
          {props.isEditor && (
            <button className={styles.addBtnFooter} onClick={openAddMessage}>+ Add</button>
          )}
        </div>

      </div>

      {renderModals()}
    </div>
  );
};

export default LeadershipMessagePanel;
