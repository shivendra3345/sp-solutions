import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import styles from '../news/NewsList.module.scss';

interface IFeedItem {
    source: 'news' | 'announcement';
    Id: number;
    Title: string;
    Created: string;
    Description?: string;
    FileRef?: string;
    AbsoluteUrl?: string;
    AttachmentUrl?: string;
    CanvasContent1?: string;
    Author?: string;
}

const CombinedFeed: React.FC<{ context: any; maxItems?: number }> = ({ context, maxItems = 6 }) => {
    const [items, setItems] = useState<IFeedItem[]>([]);
    const [selected, setSelected] = useState<IFeedItem | null>(null);
    const MODAL_HOST_ID = 'sp-solutions-modal-host';

    // ensure modal host exists (HeroBanner also creates it, but be defensive)
    React.useEffect(() => {
        if (typeof document === 'undefined') return;
        let host = document.getElementById(MODAL_HOST_ID);
        if (!host) {
            host = document.createElement('div');
            host.id = MODAL_HOST_ID;
            document.body.appendChild(host);
        }
    }, []);

    useEffect(() => {
        const fetchAll = async () => {
            try {
                // News (sitepages/pages promoted state)
                const newsUrl = `${context.pageContext.web.absoluteUrl}/_api/sitepages/pages?$filter=PromotedState eq 2&$select=Id,Title,Created,FileRef,AbsoluteUrl,Description,CanvasContent1,Author/Title&$expand=Author&$orderby=Created desc&$top=${maxItems}`;
                const annUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Announcement')/items?$select=Id,Title,Content,Created,AttachmentFiles,Author/Title&$expand=AttachmentFiles,Author&$orderby=Created desc&$top=${maxItems}`;

                const [newsRes, annRes] = await Promise.all([
                    context.spHttpClient.get(newsUrl, SPHttpClient.configurations.v1),
                    context.spHttpClient.get(annUrl, SPHttpClient.configurations.v1)
                ]);

                const newsJson = await newsRes.json();
                const annJson = await annRes.json();

                const getAuthor = (obj: any): string | undefined => {
                    if (!obj) return undefined;
                    if (obj.Author && typeof obj.Author === 'object' && obj.Author.Title) return obj.Author.Title;
                    if (obj.Author && typeof obj.Author === 'string') {
                        // sometimes SharePoint returns '1;#Firstname Lastname'
                        const m = obj.Author.match(/#(.*)$/);
                        if (m && m[1]) return m[1];
                        return obj.Author;
                    }
                    if (obj.FieldValuesAsText && obj.FieldValuesAsText.Author) return obj.FieldValuesAsText.Author;
                    if (obj['AuthorId']) return String(obj['AuthorId']);
                    if (obj.CreatedBy && obj.CreatedBy.Title) return obj.CreatedBy.Title;
                    return undefined;
                };

                const newsItems: IFeedItem[] = (newsJson.value || newsJson.d?.results || []).map((n: any) => ({
                    source: 'news',
                    Id: n.Id,
                    Title: n.Title,
                    Created: n.Created,
                    Description: n.Description,
                    FileRef: n.FileRef,
                    AbsoluteUrl: n.AbsoluteUrl
                    ,
                    CanvasContent1: n.CanvasContent1,
                    Author: getAuthor(n)
                }));

                const annItems: IFeedItem[] = (annJson.value || annJson.d?.results || []).map((a: any) => ({
                    source: 'announcement',
                    Id: a.Id,
                    Title: a.Title,
                    Created: a.Created,
                    Description: a.Content,
                    AttachmentUrl: (a.AttachmentFiles && a.AttachmentFiles.length > 0) ? a.AttachmentFiles[0].ServerRelativeUrl : undefined,
                    Author: getAuthor(a)
                }));

                const merged = [...newsItems, ...annItems]
                    .sort((a, b) => (new Date(b.Created).getTime() - new Date(a.Created).getTime()))
                    .slice(0, maxItems);

                setItems(merged);
            } catch (e) {
                console.error('Error fetching combined feed:', e);
                setItems([]);
            }
        };

        fetchAll();
    }, [context, maxItems]);

    const handleClick = (it: IFeedItem) => setSelected(it);

    const extractThumbnail = (item: IFeedItem): string | undefined => {
        if (!item.CanvasContent1) return undefined;
        try {
            const content = JSON.parse(item.CanvasContent1);
            const imagePart = Array.isArray(content)
                ? content.find(
                    (c: any) =>
                        c?.controlType === 3 &&
                        c?.webPartData?.serverProcessedContent?.imageSources
                )
                : undefined;
            const imgSources = imagePart?.webPartData?.serverProcessedContent?.imageSources;
            let url: any = undefined;
            if (!imgSources) return undefined;
            if (typeof imgSources === 'string') url = imgSources;
            else if (imgSources.imageSource) url = imgSources.imageSource;
            else if (typeof imgSources === 'object') {
                const first: any = Object.values(imgSources)[0];
                if (typeof first === 'string') url = first;
                else if (first && (first as any).imageSource) url = (first as any).imageSource;
            }

            const normalize = (u?: string) => {
                if (!u) return u;
                if (u.indexOf('~sitecollection') === 0) return u.replace('~sitecollection', context.pageContext.web.absoluteUrl);
                if (u.indexOf('~site') === 0) return u.replace('~site', context.pageContext.web.absoluteUrl);
                if (u.startsWith('/')) return `${window.location.origin}${u}`;
                return u;
            };

            return normalize(url);
        } catch { return undefined; }
    };

    // We no longer inject parsed CanvasContent1 HTML into the modal.
    // News will be displayed using an iframe (when available) to keep rendering consistent with the original page.

    return (
        <div className={styles.newsListInline}>
            {items.length === 0 ? (
                <div className={styles.noItems}>No items found.</div>
            ) : (
                items.map(item => (
                    <div
                        key={`${item.source}-${item.Id}`}
                        className={styles.newsCard}
                        role="button"
                        tabIndex={0}
                        onClick={() => handleClick(item)}
                        onKeyDown={(e) => e.key === 'Enter' && handleClick(item)}
                    >
                        {item.source === 'announcement' && item.AttachmentUrl && (
                            <img src={item.AttachmentUrl} alt={item.Title} className={styles.newsImage} loading="lazy" />
                        )}
                        {item.source === 'news' && extractThumbnail(item) && (
                            <img src={extractThumbnail(item)!} alt={item.Title} className={styles.newsImage} loading="lazy" />
                        )}

                        <h3 className={styles.newsTitle}>{item.Title}</h3>
                        {item.Description && (
                            <p className={styles.newsExcerpt}>
                                {item.Description.length > 120 ? item.Description.substring(0, 120) + '…' : item.Description}
                            </p>
                        )}
                        <p className={styles.newsMeta}>{new Date(item.Created).toLocaleDateString()}</p>
                    </div>
                ))
            )}

            <Dialog
                hidden={!selected}
                onDismiss={() => setSelected(null)}
                dialogContentProps={{
                    type: DialogType.normal,
                    title: selected ? selected.Title : undefined,
                    subText: selected ? new Date(selected.Created).toLocaleDateString() : undefined
                }}
                modalProps={{
                    isBlocking: false,
                    layerProps: { hostId: MODAL_HOST_ID }
                }}
            >
                {selected && (
                    <div>
                        <button className={styles.closeButton} onClick={() => setSelected(null)} aria-label="Close">×</button>
                        {selected.source === 'news' && (
                            <>
                                {selected.FileRef || selected.AbsoluteUrl ? (
                                    <iframe
                                        src={`${selected.FileRef || selected.AbsoluteUrl}${(selected.FileRef || selected.AbsoluteUrl)!.includes('?') ? '&' : '?'}chromeless=1`}
                                        className={styles.newsIframe}
                                        title={selected.Title}
                                    />
                                ) : (
                                    <div className={styles.modalBody}>
                                        <p>Unable to embed this news page. <a href={selected.FileRef || selected.AbsoluteUrl} target="_blank" rel="noreferrer">Open in new tab</a></p>
                                    </div>
                                )}
                            </>
                        )}

                        {selected.source === 'announcement' && selected.AttachmentUrl && (
                            <img src={selected.AttachmentUrl} alt={selected.Title} className={styles.modalImage} />
                        )}

                        {selected.source === 'announcement' && selected.Description && (
                            <div className={styles.modalBody}><p>{selected.Description}</p></div>
                        )}

                        <DialogFooter>
                            <PrimaryButton
                                text="Open page"
                                onClick={() => {
                                    const href = selected.source === 'news'
                                        ? (selected.FileRef || selected.AbsoluteUrl)
                                        : (selected.AttachmentUrl || selected.FileRef || selected.AbsoluteUrl);
                                    if (href) window.open(href, '_blank');
                                }}
                            />
                            <DefaultButton text="Close" onClick={() => setSelected(null)} />
                        </DialogFooter>
                    </div>
                )}
            </Dialog>
        </div>
    );
};

export default CombinedFeed;
