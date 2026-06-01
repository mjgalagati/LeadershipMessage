import * as React from "react";
import * as ReactDOM from "react-dom";
import { useState, useEffect, useRef } from "react";
import { Icon } from "@fluentui/react";
import { ILeadershipMessage } from "../models/ILeadershipMessage";
import { IExecutive } from "../models/IExecutive";
import styles from "./AddEditMessageModal.module.scss";
import { WebPartContext } from "@microsoft/sp-webpart-base";

interface AddEditMessageModalProps {
  isOpen: boolean;
  mode: "add" | "edit";
  message?: ILeadershipMessage;
  executives: IExecutive[];
  context: WebPartContext;
  onDismiss: () => void;
  onSave: (data: Partial<ILeadershipMessage>, attachments?: File[], deletedAttachmentNames?: string[]) => Promise<void>;
}

const AddEditMessageModal: React.FC<AddEditMessageModalProps> = ({
  isOpen, mode, message, executives, onDismiss, onSave,
}) => {
  const [title, setTitle] = useState("");
  const [selectedExecId, setSelectedExecId] = useState<number | undefined>(undefined);
  const [newFiles, setNewFiles] = useState<File[]>([]);
  const [deletedNames, setDeletedNames] = useState<string[]>([]);
  const [existingAttachments, setExistingAttachments] = useState<{ FileName: string; ServerRelativeUrl: string }[]>([]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const bodyRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const selectedExec = executives.find(e => e.Id === selectedExecId);

  useEffect(() => {
    if (!isOpen) return;
    setTitle(message?.Title ?? "");
    setSelectedExecId(message?.Executive?.Id);
    setExistingAttachments(message?.Attachments ?? []);
    setNewFiles([]);
    setDeletedNames([]);
    setError(undefined);
    setSaving(false);
    if (bodyRef.current) {
      bodyRef.current.innerHTML = message?.Body ?? "";
    }
  }, [isOpen, message]);

  if (!isOpen) return <></>;

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const handleFileAdd = (e: React.ChangeEvent<HTMLInputElement>): void => {
    if (e.target.files) {
      setNewFiles(prev => [...prev, ...Array.from(e.target.files as FileList)]);
    }
    e.target.value = "";
  };

  const removeNewFile = (idx: number): void => {
    setNewFiles(prev => prev.filter((_, i) => i !== idx));
  };

  const removeExistingAttachment = (fileName: string): void => {
    setExistingAttachments(prev => prev.filter(a => a.FileName !== fileName));
    setDeletedNames(prev => [...prev, fileName]);
  };

  const handleSave = async (): Promise<void> => {
    if (!title.trim()) { setError("Title is required."); return; }
    setSaving(true);
    setError(undefined);
    try {
      const data: Partial<ILeadershipMessage> = {
        Title: title.trim(),
        Body: bodyRef.current?.innerHTML ?? "",
        Executive: selectedExec,
      };
      await onSave(data, newFiles.length > 0 ? newFiles : undefined, deletedNames.length > 0 ? deletedNames : undefined);
    } catch {
      setError("Failed to save. Please try again.");
      setSaving(false);
    }
  };

  return ReactDOM.createPortal(
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        <div className={styles.modalHeader}>
          <h2 className={styles.modalTitle}>{mode === "add" ? "Add Message" : "Edit Message"}</h2>
          <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
            <Icon iconName="Cancel" />
          </button>
        </div>

        <div className={styles.scrollBody}>
          <div className={styles.form}>

            {/* ── Executive picker ── */}
            <div className={styles.field}>
              <label className={styles.label}>Executive</label>
              <div className={styles.execPickerRow}>
                <select
                  className={styles.select}
                  value={selectedExecId ?? ""}
                  onChange={e => setSelectedExecId(e.target.value ? Number(e.target.value) : undefined)}
                >
                  <option value="">— None —</option>
                  {executives.map(exec => (
                    <option key={exec.Id} value={exec.Id}>{exec.Title} — {exec.Position}</option>
                  ))}
                </select>
                {selectedExec && (
                  <div className={styles.execPreview}>
                    {selectedExec.PhotoUrl ? (
                      <img
                        src={selectedExec.PhotoUrl}
                        alt={selectedExec.Title}
                        className={styles.execPreviewAvatar}
                        style={{ objectPosition: `${selectedExec.PhotoFocalX}% ${selectedExec.PhotoFocalY}%` }}
                      />
                    ) : (
                      <div className={styles.execPreviewFallback}>
                        {selectedExec.Title.charAt(0).toUpperCase()}
                      </div>
                    )}
                    <div className={styles.execPreviewInfo}>
                      <span className={styles.execPreviewName}>{selectedExec.Title}</span>
                      <span className={styles.execPreviewPos}>{selectedExec.Position}</span>
                    </div>
                  </div>
                )}
              </div>
            </div>

            {/* ── Title ── */}
            <div className={styles.field}>
              <label className={styles.label}>Title <span className={styles.required}>*</span></label>
              <input
                type="text"
                className={styles.input}
                value={title}
                onChange={e => setTitle(e.target.value)}
                placeholder="Message title"
              />
            </div>

            {/* ── Body (contentEditable) ── */}
            <div className={styles.field}>
              <label className={styles.label}>Body</label>
              <div className={styles.bodyToolbar}>
                <button type="button" onMouseDown={e => { e.preventDefault(); document.execCommand("bold"); }} className={styles.toolbarBtn} title="Bold"><b>B</b></button>
                <button type="button" onMouseDown={e => { e.preventDefault(); document.execCommand("italic"); }} className={styles.toolbarBtn} title="Italic"><i>I</i></button>
                <button type="button" onMouseDown={e => { e.preventDefault(); document.execCommand("underline"); }} className={styles.toolbarBtn} title="Underline"><u>U</u></button>
                <button type="button" onMouseDown={e => { e.preventDefault(); document.execCommand("insertUnorderedList"); }} className={styles.toolbarBtn} title="Bullet list"><Icon iconName="BulletedList" /></button>
                <button type="button" onMouseDown={e => { e.preventDefault(); document.execCommand("insertOrderedList"); }} className={styles.toolbarBtn} title="Numbered list"><Icon iconName="NumberedList" /></button>
              </div>
              <div
                ref={bodyRef}
                className={styles.bodyEditor}
                contentEditable
                suppressContentEditableWarning
                data-placeholder="Write your message here…"
              />
            </div>

            {/* ── Attachments ── */}
            <div className={styles.field}>
              <label className={styles.label}>Attachments</label>

              {existingAttachments.length > 0 && (
                <div className={styles.attachmentList}>
                  {existingAttachments.map(att => (
                    <div key={att.FileName} className={styles.attachmentRow}>
                      <Icon iconName="Attach" className={styles.attachIcon} />
                      <span className={styles.attachName}>{att.FileName}</span>
                      <button type="button" className={styles.removeBtn} onClick={() => removeExistingAttachment(att.FileName)} aria-label="Remove">
                        <Icon iconName="Cancel" />
                      </button>
                    </div>
                  ))}
                </div>
              )}

              {newFiles.length > 0 && (
                <div className={styles.attachmentList}>
                  {newFiles.map((f, idx) => (
                    <div key={idx} className={styles.attachmentRow}>
                      <Icon iconName="Attach" className={styles.attachIcon} />
                      <span className={styles.attachName}>{f.name}</span>
                      <button type="button" className={styles.removeBtn} onClick={() => removeNewFile(idx)} aria-label="Remove">
                        <Icon iconName="Cancel" />
                      </button>
                    </div>
                  ))}
                </div>
              )}

              <input
                ref={fileInputRef}
                type="file"
                multiple
                className={styles.fileInputHidden}
                onChange={handleFileAdd}
              />
              <button type="button" className={styles.uploadBtn} onClick={() => fileInputRef.current?.click()}>
                <Icon iconName="Upload" /> Attach files
              </button>
            </div>

            {error && <p className={styles.errorMsg}>{error}</p>}
          </div>
        </div>

        <div className={styles.footer}>
          <button className={styles.cancelBtn} onClick={onDismiss} disabled={saving}>Cancel</button>
          <button className={styles.saveBtn} onClick={handleSave} disabled={saving}>
            {saving ? "Saving…" : "Save"}
          </button>
        </div>

      </div>
    </div>,
    document.body
  );
};

export default AddEditMessageModal;
