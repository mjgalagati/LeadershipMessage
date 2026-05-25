import * as React from 'react';
import { PrimaryButton, Icon, SearchBox, DatePicker, IconButton } from '@fluentui/react';
import { IAlert } from '../models/IAlert';
import styles from './ViewAllPanel.module.scss';
import { useState } from 'react';

interface ViewAllPanelProps {
  alerts: IAlert[];
  isOpen: boolean;
  onDismiss: () => void;
  onSelectAlert: (alert: IAlert) => void;
  onAddAlert?: () => void;
}

const ViewAllPanel: React.FC<ViewAllPanelProps> = ({
  alerts, isOpen, onDismiss, onSelectAlert, onAddAlert,
}): JSX.Element => {
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [dateFrom, setDateFrom] = useState<Date | undefined>(undefined);
  const [dateTo, setDateTo] = useState<Date | undefined>(undefined);
  const [showDateFilter, setShowDateFilter] = useState<boolean>(false);

  if (!isOpen) return <></>;

  const filteredAlerts = alerts.filter(alert => {
    const matchesSearch =
      alert.Title.toLowerCase().includes(searchQuery.toLowerCase()) ||
      alert.AlertType.toLowerCase().includes(searchQuery.toLowerCase()) ||
      alert.Body?.toLowerCase().includes(searchQuery.toLowerCase());

    let matchesDateRange = true;
    if (dateFrom || dateTo) {
      const publishDate = new Date(alert.PublishDate);
      if (dateFrom && dateTo) matchesDateRange = publishDate >= dateFrom && publishDate <= dateTo;
      else if (dateFrom) matchesDateRange = publishDate >= dateFrom;
      else if (dateTo) matchesDateRange = publishDate <= dateTo;
    }

    return matchesSearch && matchesDateRange;
  });

  const criticalAlerts    = filteredAlerts.filter(a => a.Priority?.trim() === 'Critical' && a.Status?.trim() === 'Active');
  const highAlerts        = filteredAlerts.filter(a => a.Priority?.trim() === 'High' && a.Status?.trim() === 'Active');
  const activeAlerts      = filteredAlerts.filter(a => (a.Priority?.trim() === 'Medium' || a.Priority?.trim() === 'Low') && a.Status?.trim() === 'Active');
  const scheduledAlerts   = filteredAlerts.filter(a => a.Status?.trim() === 'Scheduled');
  const draftAlerts       = filteredAlerts.filter(a => a.Status?.trim() === 'Draft');
  const expiredAlerts     = filteredAlerts.filter(a => a.Status?.trim() === 'Expired');

  const formatDate = (date?: Date): string => {
    if (!date) return 'N/A';
    return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
  };

  const getStatusClass = (status: string): string => {
    if (status === 'Active')    return styles.statusActive;
    if (status === 'Scheduled') return styles.statusScheduled;
    if (status === 'Expired')   return styles.statusExpired;
    return styles.statusDraft;
  };

  const getPriorityClass = (priority: string): string => {
    if (priority === 'Critical') return styles.priorityCritical;
    if (priority === 'High')     return styles.priorityHigh;
    if (priority === 'Medium')   return styles.priorityMedium;
    return styles.priorityLow;
  };

  const handleBackdropClick = (e: React.MouseEvent<HTMLDivElement>): void => {
    if (e.target === e.currentTarget) onDismiss();
  };

  const renderAlertCard = (alert: IAlert): JSX.Element => (
    <div
      key={alert.Id}
      className={styles.alertCard}
      onClick={() => onSelectAlert(alert)}
    >
      {alert.BannerImageUrl && (
        <div className={styles.cardBanner}>
          <img src={alert.BannerImageUrl} alt={alert.Title} />
        </div>
      )}
      <div className={styles.cardContent}>
        <div className={styles.cardHeader}>
          <h3 className={styles.cardTitle}>{alert.Title}</h3>
          <div className={styles.cardBadges}>
            <span className={`${styles.priorityBadge} ${getPriorityClass(alert.Priority)}`}>
              {alert.Priority}
            </span>
            <span className={`${styles.statusBadge} ${getStatusClass(alert.Status)}`}>
              {alert.Status}
            </span>
          </div>
        </div>
        <div className={styles.cardInfo}>
          <div className={styles.infoItem}>
            <Icon iconName="Calendar" className={styles.infoIcon} />
            <span>{formatDate(alert.PublishDate)}</span>
          </div>
          {alert.ExpiryDate && (
            <div className={styles.infoItem}>
              <Icon iconName="EventDateMissed12" className={styles.infoIcon} />
              <span>Expires: {formatDate(alert.ExpiryDate)}</span>
            </div>
          )}
          <div className={styles.infoItem}>
            <Icon iconName="Tag" className={styles.infoIcon} />
            <span className={styles.typeTag}>{alert.AlertType}</span>
          </div>
          {alert.TargetAudienceType && alert.TargetAudienceType !== 'All' && (
            <div className={styles.infoItem}>
              <Icon iconName="People" className={styles.infoIcon} />
              <span>{alert.TargetAudienceType === 'Specific' ? 'Specific Audience' : 'All Except'}</span>
            </div>
          )}
        </div>
      </div>
    </div>
  );

  return (
    <div className={styles.backdrop} onClick={handleBackdropClick} role="dialog" aria-modal="true">
      <div className={styles.modal} onClick={e => e.stopPropagation()}>

        {/* ── Header ── */}
        <div className={styles.panelHeader}>
          <div className={styles.headerTop}>
            <h2 className={styles.panelTitle}>All Alerts</h2>
            <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
              {onAddAlert && (
                <PrimaryButton
                  text="Add Alert"
                  iconProps={{ iconName: 'Add' }}
                  onClick={onAddAlert}
                  className={styles.addButton}
                />
              )}
              <button className={styles.closeBtn} onClick={onDismiss} aria-label="Close">
                <Icon iconName="Cancel" />
              </button>
            </div>
          </div>

          <div className={styles.filterBar}>
            <SearchBox
              placeholder="Search alerts..."
              onChange={(_, newValue) => setSearchQuery(newValue || '')}
              className={styles.searchBox}
            />
            <IconButton
              iconProps={{ iconName: showDateFilter ? 'ChevronUp' : 'Calendar' }}
              onClick={() => setShowDateFilter(!showDateFilter)}
              className={styles.dateFilterToggle}
              toggle
              checked={showDateFilter}
            />
          </div>

          {showDateFilter && (
            <div className={styles.dateFilters}>
              <div className={styles.datePickerWrapper}>
                <label className={styles.dateLabel}>From:</label>
                <DatePicker
                  placeholder="Select start date"
                  value={dateFrom}
                  onSelectDate={(date) => setDateFrom(date || undefined)}
                  formatDate={(date) => date?.toLocaleDateString() || ''}
                  className={styles.datePicker}
                />
                {dateFrom && (
                  <IconButton
                    iconProps={{ iconName: 'Clear' }}
                    onClick={() => setDateFrom(undefined)}
                    className={styles.clearButton}
                  />
                )}
              </div>
              <div className={styles.datePickerWrapper}>
                <label className={styles.dateLabel}>To:</label>
                <DatePicker
                  placeholder="Select end date"
                  value={dateTo}
                  onSelectDate={(date) => setDateTo(date || undefined)}
                  formatDate={(date) => date?.toLocaleDateString() || ''}
                  minDate={dateFrom}
                  className={styles.datePicker}
                />
                {dateTo && (
                  <IconButton
                    iconProps={{ iconName: 'Clear' }}
                    onClick={() => setDateTo(undefined)}
                    className={styles.clearButton}
                  />
                )}
              </div>
              {(dateFrom || dateTo) && (
                <IconButton
                  iconProps={{ iconName: 'ClearFilter' }}
                  onClick={() => { setDateFrom(undefined); setDateTo(undefined); }}
                  className={styles.clearAllButton}
                />
              )}
            </div>
          )}

          <div className={styles.alertCount}>
            <span>{filteredAlerts.length} alert{filteredAlerts.length !== 1 ? 's' : ''} found</span>
          </div>
        </div>

        {/* ── Scrollable Content ── */}
        <div className={styles.panelContent}>
          {filteredAlerts.length === 0 && (
            <div className={styles.emptyState}>
              <Icon iconName="Megaphone" className={styles.emptyIcon} />
              <p>No alerts found</p>
              <span>Try adjusting your search or add a new alert</span>
            </div>
          )}

          {criticalAlerts.length > 0 && (
            <div className={styles.alertSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Warning" className={styles.sectionIconCritical} />
                <h3>Critical ({criticalAlerts.length})</h3>
              </div>
              <div className={styles.alertGrid}>
                {criticalAlerts.map(renderAlertCard)}
              </div>
            </div>
          )}

          {highAlerts.length > 0 && (
            <div className={styles.alertSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="SortUp" className={styles.sectionIconHigh} />
                <h3>High Priority ({highAlerts.length})</h3>
              </div>
              <div className={styles.alertGrid}>
                {highAlerts.map(renderAlertCard)}
              </div>
            </div>
          )}

          {activeAlerts.length > 0 && (
            <div className={styles.alertSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="MegaphoneSolid" className={styles.sectionIcon} />
                <h3>Active ({activeAlerts.length})</h3>
              </div>
              <div className={styles.alertGrid}>
                {activeAlerts.map(renderAlertCard)}
              </div>
            </div>
          )}

          {scheduledAlerts.length > 0 && (
            <div className={styles.alertSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="ScheduleEventAction" className={styles.sectionIcon} />
                <h3>Scheduled ({scheduledAlerts.length})</h3>
              </div>
              <div className={styles.alertGrid}>
                {scheduledAlerts.map(renderAlertCard)}
              </div>
            </div>
          )}

          {draftAlerts.length > 0 && (
            <div className={styles.alertSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="Edit" className={styles.sectionIcon} />
                <h3>Draft ({draftAlerts.length})</h3>
              </div>
              <div className={styles.alertGrid}>
                {draftAlerts.map(renderAlertCard)}
              </div>
            </div>
          )}

          {expiredAlerts.length > 0 && (
            <div className={styles.alertSection}>
              <div className={styles.sectionHeader}>
                <Icon iconName="History" className={styles.sectionIcon} />
                <h3>Expired ({expiredAlerts.length})</h3>
              </div>
              <div className={styles.alertGrid}>
                {expiredAlerts.map(renderAlertCard)}
              </div>
            </div>
          )}
        </div>

      </div>
    </div>
  );
};

export default ViewAllPanel;
