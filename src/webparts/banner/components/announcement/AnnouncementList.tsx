import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './AnnouncementList.module.scss';

interface IAnnouncement {
    Id: number;
    Title: string;
    Content?: string;
    Created: string;
    AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
}

interface IAnnouncementResponse {
    value?: IAnnouncement[];
    d?: { results?: IAnnouncement[] };
}

const AnnouncementList: React.FC<{ context: any }> = ({ context }) => {
    const [announcements, setAnnouncements] = useState<IAnnouncement[]>([]);
    const [selected, setSelected] = useState<IAnnouncement | null>(null);

    useEffect(() => {
        const url =
            `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Announcement')/items` +
            `?$select=Id,Title,Content,Created,AttachmentFiles&$expand=AttachmentFiles&$orderby=Created desc&$top=6`;

        context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((res: SPHttpClientResponse) => res.json())
            .then((data: IAnnouncementResponse) => {
                const items = data.value || data.d?.results || [];
                setAnnouncements(items);
            })
            .catch((err: any) => console.error("Error fetching announcements:", err));
    }, [context]);

    const getImageUrl = (item: IAnnouncement): string => {
        if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
            return item.AttachmentFiles[0].ServerRelativeUrl;
        }
        // fallback default image
        return `${context.pageContext.web.absoluteUrl}/SiteAssets/images/default-announcement.jpg`;
    };

    return (
        <>
            {announcements.length === 0 ? (
                <div className={styles.noItems}>No announcements found.</div>
            ) : (
                announcements.map(item => (
                    <div
                        key={item.Id}
                        className={styles.announcementCard}
                        role="button"
                        tabIndex={0}
                        onClick={() => setSelected(item)}
                        onKeyDown={(e) => e.key === 'Enter' && setSelected(item)}
                    >
                        <img
                            src={getImageUrl(item)}
                            alt={item.Title}
                            className={styles.announcementImage}
                            loading="lazy"
                        />
                        <h3 className={styles.announcementTitle}>{item.Title}</h3>
                        {item.Content && (
                            <p className={styles.announcementExcerpt}>
                                {item.Content.length > 120
                                    ? item.Content.substring(0, 120) + '…'
                                    : item.Content}
                            </p>
                        )}
                        <p className={styles.announcementMeta}>
                            {new Date(item.Created).toLocaleDateString()}
                        </p>
                    </div>
                ))
            )}

            {selected && (
                <div className={styles.modalOverlay} onClick={() => setSelected(null)}>
                    <div className={styles.modalContent} onClick={(e) => e.stopPropagation()}>
                        <button
                            className={styles.closeButton}
                            onClick={() => setSelected(null)}
                            aria-label="Close"
                        >
                            ×
                        </button>

                        <div className={styles.modalHeader}>
                            <h2 className={styles.modalTitle}>{selected.Title}</h2>
                            <p className={styles.modalMeta}>
                                {new Date(selected.Created).toLocaleDateString()}
                            </p>
                        </div>

                        <img
                            src={getImageUrl(selected)}
                            alt={selected.Title}
                            className={styles.modalImage}
                        />

                        {selected.Content && (
                            <div className={styles.modalBody}>
                                <p>{selected.Content}</p>
                            </div>
                        )}
                    </div>
                </div>
            )}
        </>
    );
};

export default AnnouncementList;