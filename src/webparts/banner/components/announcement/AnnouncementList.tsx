import * as React from 'react';
import { useEffect, useState } from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { SPHttpClient } from '@microsoft/sp-http';
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
        const fetchAnnouncements = async () => {
            const url =
                `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Announcement')/items` +
                `?$select=Id,Title,Content,Created,AttachmentFiles,Author/Title,AuthorId,FieldValuesAsText&$expand=AttachmentFiles,Author&$orderby=Created desc&$top=6`;
            try {
                const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
                const data: IAnnouncementResponse = await res.json();
                console.debug('AnnouncementList: raw response', data);
                const items: any[] = data.value || data.d?.results || [];

                // For items missing an expanded Author, try a fallback lookup by AuthorId
                const processed = await Promise.all(items.map(async (it: any) => {
                    try {
                        if (!getAuthorName(it)) {
                            // try FieldValuesAsText which sometimes contains readable author text
                            if (it.FieldValuesAsText && it.FieldValuesAsText.Author) {
                                it.__AuthorName = it.FieldValuesAsText.Author;
                            }

                            const aid = it.AuthorId || it['AuthorId'] || it.AuthorId0;
                            if (!it.__AuthorName && aid) {
                                try {
                                    const ur = await context.spHttpClient.get(`${context.pageContext.web.absoluteUrl}/_api/web/getuserbyid(${aid})?$select=Title`, SPHttpClient.configurations.v1);
                                    if (ur && ur.ok) {
                                        const uj = await ur.json();
                                        it.__AuthorName = uj && (uj.Title || uj.Title) || undefined;
                                    }
                                } catch (ux) {
                                    console.debug('AnnouncementList: getuserbyid failed for', aid, ux);
                                }
                            }

                            // final fallback: fetch single item expanded with Author (some tenants omit lookup fields in list queries)
                            if (!it.__AuthorName) {
                                try {
                                    const singleUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Announcement')/items(${it.Id})?$select=Author/Title&$expand=Author`;
                                    const sr = await context.spHttpClient.get(singleUrl, SPHttpClient.configurations.v1);
                                    if (sr && sr.ok) {
                                        const sj = await sr.json();
                                        if (sj && sj.Author && sj.Author.Title) it.__AuthorName = sj.Author.Title;
                                    }
                                } catch (sx) {
                                    console.debug('AnnouncementList: single item author fetch failed for', it.Id, sx);
                                }
                            }
                        }
                    } catch (inner) {
                        /* ignore per-item author lookup errors */
                    }
                    return it;
                }));

                setAnnouncements(processed);
            } catch (err) {
                console.error('Error fetching announcements:', err);
            }
        };

        fetchAnnouncements();
    }, [context]);

    const getImageUrl = (item: IAnnouncement): string => {
        if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
            return item.AttachmentFiles[0].ServerRelativeUrl;
        }
        // fallback default image
        return `${context.pageContext.web.absoluteUrl}/SiteAssets/images/default-announcement.jpg`;
    };

    const getAuthorName = (item: any): string | undefined => {
        if (!item) return undefined;
        if (item.Author && typeof item.Author === 'object' && item.Author.Title) return item.Author.Title;
        if (item.Author && typeof item.Author === 'string') {
            const m = item.Author.match(/#(.*)$/);
            if (m && m[1]) return m[1];
            return item.Author;
        }
        if (item.FieldValuesAsText && item.FieldValuesAsText.Author) return item.FieldValuesAsText.Author;
        if (item['AuthorId']) return String(item['AuthorId']);
        return undefined;
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
                        <div className={styles.imageWrap}>
                            <img
                                src={getImageUrl(item)}
                                alt={item.Title}
                                className={styles.announcementImage}
                                loading="lazy"
                            />
                            <div className={styles.cardOverlay} aria-hidden="true">
                                <div className={styles.overlayTitle}>{item.Title}</div>
                                <div className={styles.overlayMeta}>{getAuthorName(item) || new Date(item.Created).toLocaleDateString()}</div>
                            </div>
                        </div>
                        <h3 className={styles.announcementTitle}>{item.Title}</h3>
                        {item.Content && (
                            <p className={styles.announcementExcerpt}>
                                {item.Content.length > 120
                                    ? item.Content.substring(0, 120) + '…'
                                    : item.Content}
                            </p>
                        )}
                        <p className={styles.announcementMeta}>
                            {getAuthorName(item) ? `${getAuthorName(item)} • ${new Date(item.Created).toLocaleDateString()}` : new Date(item.Created).toLocaleDateString()}
                        </p>
                    </div>
                ))
            )}

            {selected && (
                <Dialog
                    hidden={!selected}
                    onDismiss={() => setSelected(null)}
                    dialogContentProps={{
                        type: DialogType.largeHeader,
                        title: selected ? selected.Title : 'Announcement',
                        subText: selected ? new Date(selected.Created).toLocaleDateString() : undefined
                    }}
                    modalProps={{ isBlocking: false, layerProps: { hostId: 'sp-solutions-modal-host' } }}
                >
                    <div>
                        <button className={styles.closeButton} onClick={() => setSelected(null)} aria-label="Close">×</button>
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
                    <DialogFooter>
                        <PrimaryButton onClick={() => {
                            if (!selected) return;
                            const img = getImageUrl(selected);
                            if (img) window.open(img, '_blank');
                        }} text="Open image" />
                        <DefaultButton onClick={() => setSelected(null)} text="Close" />
                    </DialogFooter>
                </Dialog>
            )}
        </>
    );
};

export default AnnouncementList;