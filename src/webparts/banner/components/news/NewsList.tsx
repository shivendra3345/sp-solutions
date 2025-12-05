import * as React from 'react';
import { useEffect, useState, useMemo } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './NewsList.module.scss';

interface IAuthor {
    Title?: string;
}

interface INewsItem {
    Id: number;
    Title: string;
    Created: string;
    FileRef?: string;        // Page URL (sometimes AbsoluteUrl instead)
    AbsoluteUrl?: string;    // fallback if FileRef not present
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
            `${context.pageContext.web.absoluteUrl}` +
            `/_api/sitepages/pages` +
            `?$filter=PromotedState eq 2` +
            `&$orderby=Created desc` +
            `&$top=10` +
            `&$select=Id,Title,Created,FileRef,AbsoluteUrl,Author/Title,Description,CanvasContent1` +
            `&$expand=Author`;

        context.spHttpClient
            .get(url, SPHttpClient.configurations.v1)
            .then((res: SPHttpClientResponse) => res.json())
            .then((data: INewsApiResponse) => {
                console.log("API response:", data);
                const items = data.value || data.d?.results || [];
                setNews(items);
            })
            .catch((err: Error) => {
                console.error('Error fetching news:', err);
                setNews([]);
            });
    }, [context]);

    // Extract thumbnail from CanvasContent1
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
        } catch {
            return undefined;
        }
    };

    // Build iframe URL with chromeless
    const iframeSrc = useMemo(() => {
        if (!selectedNews) return '';
        const pageUrl = selectedNews.FileRef || selectedNews.AbsoluteUrl;
        if (!pageUrl) return '';
        const hasQuery = pageUrl.includes('?');
        const sep = hasQuery ? '&' : '?';
        return `${pageUrl}${sep}chromeless=1`;
    }, [selectedNews]);

    return (
        <div className={styles.newsCollage}>
            <h2 className={styles.sectionTitle}>Latest News</h2>

            {news.length === 0 ? (
                <p className={styles.noItems}>No news found.</p>
            ) : (
                news.map((item) => (
                    <div
                        key={item.Id}
                        className={styles.newsCard}
                        onClick={() => {
                            console.log("Clicked:", item);
                            setSelectedNews(item);
                        }}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(e) => e.key === 'Enter' && setSelectedNews(item)}
                    >
                        {extractThumbnail(item) && (
                            <img
                                src={extractThumbnail(item)!}
                                alt={item.Title}
                                className={styles.newsImage}
                            />
                        )}
                        <div className={styles.newsContent}>
                            <h3 className={styles.newsTitle}>{item.Title}</h3>
                            {item.Description && (
                                <p className={styles.newsExcerpt}>
                                    {item.Description.length > 140
                                        ? item.Description.substring(0, 140) + '…'
                                        : item.Description}
                                </p>
                            )}
                            <p className={styles.newsMeta}>
                                By {item.Author?.Title || 'Unknown'} on{' '}
                                {new Date(item.Created).toLocaleDateString()}
                            </p>
                        </div>
                    </div>
                ))
            )}

            {selectedNews && iframeSrc && (
                <div className={styles.modalOverlay} onClick={() => setSelectedNews(null)}>
                    <div className={styles.modalContent} onClick={(e) => e.stopPropagation()}>
                        <button
                            className={styles.closeButton}
                            onClick={() => setSelectedNews(null)}
                            aria-label="Close"
                        >
                            ×
                        </button>
                        <h2 className={styles.modalTitle}>{selectedNews.Title}</h2>
                        <iframe
                            src={iframeSrc}
                            className={styles.newsIframe}
                            title={selectedNews.Title}
                        />
                    </div>
                </div>
            )}
        </div>
    );
};

export default NewsList;