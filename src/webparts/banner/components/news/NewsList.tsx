import * as React from 'react';
import { useEffect, useState, useMemo } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import styles from './NewsList.module.scss';

interface IAuthor { Title?: string; }
interface INewsItem {
    Id: number;
    Title: string;
    Created: string;
    FileRef?: string;
    AbsoluteUrl?: string;
    Author?: IAuthor;
    Description?: string;
    CanvasContent1?: string;
}
interface INewsApiResponse {
    value?: INewsItem[];
    d?: { results?: INewsItem[] };
}

const NewsList: React.FC<{ context: any }> = ({ context }) => {
    const [news, setNews] = useState<INewsItem[]>([]);
    const [selectedNews, setSelectedNews] = useState<INewsItem | null>(null);

    useEffect(() => {
        const url =
            `${context.pageContext.web.absoluteUrl}/_api/sitepages/pages` +
            `?$filter=PromotedState eq 2&$orderby=Created desc&$top=6` +
            `&$select=Id,Title,Created,FileRef,AbsoluteUrl,Author/Title,Description,CanvasContent1` +
            `&$expand=Author`;

        context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((res: SPHttpClientResponse) => res.json())
            .then((data: INewsApiResponse) => setNews(data.value || data.d?.results || []))
            .catch((err: any) => console.error("Error fetching news:", err));
    }, [context]);

    const extractThumbnail = (item: INewsItem): string | undefined => {
        if (!item.CanvasContent1) return undefined;
        try {
            const content = JSON.parse(item.CanvasContent1);
            const imagePart = Array.isArray(content)
                ? content.find(
                    (c: any) =>
                        c?.controlType === 3 &&
                        c?.webPartData?.serverProcessedContent?.imageSources?.imageSource
                )
                : undefined;
            return imagePart?.webPartData?.serverProcessedContent?.imageSources?.imageSource;
        } catch { return undefined; }
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

    const iframeSrc = useMemo(() => {
        if (!selectedNews) return '';
        const pageUrl = selectedNews.FileRef || selectedNews.AbsoluteUrl;
        if (!pageUrl) return '';
        const sep = pageUrl.includes('?') ? '&' : '?';
        return `${pageUrl}${sep}chromeless=1`;
    }, [selectedNews]);

    const handleCardClick = (item: INewsItem) => {
        console.log('News card clicked:', item); // debug
        setSelectedNews(item);
    };

    return (
        <div className={styles.newsListInline}>
            {news.length === 0 ? (
                <div className={styles.noItems}>No news found.</div>
            ) : (
                news.map(item => (
                    <div
                        key={item.Id}
                        className={styles.newsCard}
                        role="button"
                        tabIndex={0}
                        onClick={() => handleCardClick(item)}
                        onKeyDown={(e) => e.key === 'Enter' && handleCardClick(item)}
                    >
                        {extractThumbnail(item) && (
                            <div className={styles.imageWrap}>
                                <img
                                    src={extractThumbnail(item)!}
                                    alt={item.Title}
                                    className={styles.newsImage}
                                    loading="lazy"
                                />
                                <div className={styles.cardOverlay} aria-hidden="true">
                                    <div className={styles.overlayTitle}>{item.Title}</div>
                                    <div className={styles.overlayMeta}>By {getAuthorName(item) || 'Unknown'}</div>
                                </div>
                            </div>
                        )}
                        <h3 className={styles.newsTitle}>{item.Title}</h3>
                        {item.Description && (
                            <p className={styles.newsExcerpt}>
                                {item.Description.length > 120
                                    ? item.Description.substring(0, 120) + '…'
                                    : item.Description}
                            </p>
                        )}
                        <p className={styles.newsMeta}>
                            By {getAuthorName(item) || 'Unknown'} on {new Date(item.Created).toLocaleDateString()}
                        </p>
                    </div>
                ))
            )}

            {selectedNews && iframeSrc && (
                <Dialog
                    hidden={!selectedNews}
                    onDismiss={() => setSelectedNews(null)}
                    dialogContentProps={{
                        type: DialogType.largeHeader,
                        title: selectedNews ? selectedNews.Title : 'News',
                        subText: selectedNews ? new Date(selectedNews.Created).toLocaleDateString() : undefined
                    }}
                    modalProps={{ isBlocking: false, layerProps: { hostId: 'sp-solutions-modal-host' } }}
                >
                    <div>
                        <button className={styles.closeButton} onClick={() => setSelectedNews(null)} aria-label="Close">×</button>
                        <iframe
                            src={iframeSrc}
                            className={styles.newsIframe}
                            title={selectedNews.Title}
                        />
                    </div>
                    <DialogFooter>
                        <PrimaryButton onClick={() => {
                            if (!selectedNews) return;
                            const pageUrl = selectedNews.FileRef || selectedNews.AbsoluteUrl;
                            if (pageUrl) window.open(pageUrl, '_blank');
                        }} text="Open page" />
                        <DefaultButton onClick={() => setSelectedNews(null)} text="Close" />
                    </DialogFooter>
                </Dialog>
            )}
        </div>
    );
};

export default NewsList;