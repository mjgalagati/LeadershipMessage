import * as React from "react";
import * as ReactDOM from "react-dom";
import { useState, useEffect, useRef } from "react";
import { Icon } from "@fluentui/react";
import { IExecutive } from "../models/IExecutive";
import styles from "./AddEditExecutiveModal.module.scss";

interface AddEditExecutiveModalProps {
  isOpen: boolean;
  executive?: IExecutive;
  onDismiss: () => void;
  onSave: (data: Partial<IExecutive>, photoFile?: File) => Promise<void>;
}

const AddEditExecutiveModal: React.FC<AddEditExecutiveModalProps> = ({
  isOpen, executive, onDismiss, onSave,
}) => {
  const [name, setName] = useState("");
  const [position, setPosition] = useState("");
  const [bio, setBio] = useState("");
  const [photoUrl, setPhotoUrl] = useState("");
  const [focalX, setFocalX] = useState(50);
  const [focalY, setFocalY] = useState(50);
  const [photoFile, setPhotoFile] = useState<File | undefined>(undefined);
  const [previewSrc, setPreviewSrc] = useState<string | undefined>(undefined);
  const [dragging, setDragging] = useState(false);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState<string | undefined>(undefined);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const focalContainerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!isOpen) return;
    setName(executive?.Title ?? "");
    setPosition(executive?.Position ?? "");
    setBio(executive?.Bio ?? "");
    setPhotoUrl(executive?.PhotoUrl ?? "");
    setFocalX(executive?.PhotoFocalX ?? 50);
    setFocalY(executive?.PhotoFocalY ?? 50);
    setPhotoFile(undefined);
    setPreviewSrc(executive?.PhotoUrl || undefined);
    setError(undefined);
    setSaving(false);
  }, [isOpen, executive]);

  if (!isOpen) return <></>;

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (!file) return;
    setPhotoFile(file);
    const url = URL.createObjectURL(file);
    setPreviewSrc(url);
    setFocalX(50);
    setFocalY(50);
    e.target.value = "";
  };

  const computeFocal = (e: { clientX: number; clientY: number }): { x: number; y: number } => {
    const rect = focalContainerRef.current?.getBoundingClientRect();
    if (!rect) return { x: focalX, y: focalY };
    const x = Math.min(100, Math.max(0, Math.round(((e.clientX - rect.left) / rect.width) * 100)));
    const y = Math.min(100, Math.max(0, Math.round(((e.clientY - rect.top) / rect.height) * 100)));
    return { x, y };
  };

  const handleFocalMouseDown = (e: React.MouseEvent<HTMLDivElement>): void => {
    e.preventDefault();
    setDragging(true);
    const { x, y } = computeFocal(e);
    setFocalX(x);
    setFocalY(y);
  };

  const handleFocalMouseMove = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (!dragging) return;
    const { x, y } = computeFocal(e);
    setFocalX(x);
    setFocalY(y);
  };

  const handleFocalMouseUp = (): void => {
    setDragging(false);
  };

  const handleSave = async (): Promise<void> => {
    if (!name.trim()) { setError("Name is required."); return; }
    setSaving(true);
    setError(undefined);
    try {
      const data: Partial<IExecutive> = {
        Title: name.trim(),
        Position: position.trim(),
        Bio: bio.trim() || undefined,
        PhotoFocalX: focalX,
        PhotoFocalY: focalY,
        // PhotoUrl is set by the service after upload; pass existing url if no new file
        PhotoUrl: photoFile ? undefined : photoUrl,
      };
      await onSave(data, photoFile);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to save. Please try again.");
      setSaving(false);
    }
  };

  const hasPhoto = !!previewSrc;

  return ReactDOM.createPortal(
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        <div className={styles.modalHeader}>
          <h2 className={styles.modalTitle}>{executive ? "Edit Executive" : "Add Executive"}</h2>
          <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
            <Icon iconName="Cancel" />
          </button>
        </div>

        <div className={styles.scrollBody}>
          <div className={styles.form}>

            {/* ── Name ── */}
            <div className={styles.field}>
              <label className={styles.label}>Name <span className={styles.required}>*</span></label>
              <input
                type="text"
                className={styles.input}
                value={name}
                onChange={e => setName(e.target.value)}
                placeholder="Full name"
              />
            </div>

            {/* ── Position ── */}
            <div className={styles.field}>
              <label className={styles.label}>Position / Title</label>
              <input
                type="text"
                className={styles.input}
                value={position}
                onChange={e => setPosition(e.target.value)}
                placeholder="e.g. Chief Executive Officer"
              />
            </div>

            {/* ── Photo upload ── */}
            <div className={styles.field}>
              <label className={styles.label}>Photo</label>
              <div className={styles.photoRow}>
                <input
                  ref={fileInputRef}
                  type="file"
                  accept="image/*"
                  className={styles.fileInputHidden}
                  onChange={handleFileChange}
                />
                <button type="button" className={styles.uploadBtn} onClick={() => fileInputRef.current?.click()}>
                  <Icon iconName="Photo2" /> {hasPhoto ? "Change photo" : "Upload photo"}
                </button>
                {photoUrl && !photoFile && (
                  <span className={styles.photoUrlHint} title={photoUrl}>
                    <Icon iconName="Link" /> {photoUrl.split("/").pop()}
                  </span>
                )}
              </div>

              {/* ── Focal point tool ── */}
              {hasPhoto && (
                <div className={styles.focalSection}>
                  <p className={styles.focalHint}>
                    <Icon iconName="Crosshair" /> Drag the dot to set the focal point for cropping.
                  </p>
                  <div
                    ref={focalContainerRef}
                    className={styles.focalContainer}
                    onMouseDown={handleFocalMouseDown}
                    onMouseMove={handleFocalMouseMove}
                    onMouseUp={handleFocalMouseUp}
                    onMouseLeave={handleFocalMouseUp}
                    style={{ cursor: dragging ? "crosshair" : "pointer" }}
                  >
                    <img
                      src={previewSrc}
                      alt="Preview"
                      className={styles.focalImage}
                      draggable={false}
                      style={{ objectPosition: `${focalX}% ${focalY}%` }}
                    />
                    <div
                      className={styles.focalDot}
                      style={{ left: `${focalX}%`, top: `${focalY}%` }}
                    />
                  </div>
                  <p className={styles.focalCoords}>Focal point: {focalX}% / {focalY}%</p>
                </div>
              )}
            </div>

            {/* ── Bio ── */}
            <div className={styles.field}>
              <label className={styles.label}>Bio</label>
              <textarea
                className={styles.textarea}
                value={bio}
                onChange={e => setBio(e.target.value)}
                rows={4}
                placeholder="Short biography (optional)"
              />
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

export default AddEditExecutiveModal;
