import * as React from 'react';
import { useEffect, useState, useMemo } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './NewsList.module.scss';

interface INewsItem {
    Id: number;
    Title: string;
    Created: string;
    FileRef: string;
    Author?: { Title: string };
    Description?: string;
    CanvasContent1?: string;
}

const NewsList: React.FC<{ context: any }> = ({ context }) => {
    const [news, setNews] = useState<INewsItem[]>([]);
    const [selectedNews, setSelectedNews] = useState<INewsItem | null>(null);

    useEffect(() => {
        const url = `${context.pageContext.web.absoluteUrl}/_api/sitepages/pages?$filter=PromotedState eq 2&$orderby=Created desc&$top=5&$select=Id,Title,Created,FileRef,Author/Title,Description,CanvasContent1&$expand=Author`;
        context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((res: SPHttpClientResponse) => res.json())
            .then((data: { value?: INewsItem[]; d?: { results?: INewsItem[] } }) => {
                const items = data.value || data.d?.results || [];
                setNews(items);
            })
            .catch((err: Error) => console.error("Error fetching news:", err));
    }, [context]);

    const stripHtml = (html: string) => {
        const tmp = document.createElement("div");
        tmp.innerHTML = html;
        return tmp.textContent || tmp.innerText || "";
    };

    // Extract thumbnail image from Image web part inside CanvasContent1
    const extractImageUrl = (item: INewsItem): string | undefined => {
        if (!item.CanvasContent1) return undefined;
        try {
            const content = JSON.parse(item.CanvasContent1);
            const imagePart = content.find(
                (c: any) =>
                    c.controlType === 3 &&
                    c.webPartData?.serverProcessedContent?.imageSources?.imageSource
            );
            return imagePart?.webPartData.serverProcessedContent.imageSources.imageSource;
        } catch {
            return undefined;
        }
    };

    // Build full HTML for modal by concatenating text and image controls
    const buildFullHtml = (item: INewsItem): string => {
        if (!item.CanvasContent1) return '';
        try {
            const content = JSON.parse(item.CanvasContent1);
            const parts: string[] = [];

            for (const c of content) {
                if (c.controlType === 4 && c.innerHTML) {
                    parts.push(c.innerHTML);
                }
                if (c.controlType === 3 && c.webPartData?.serverProcessedContent?.imageSources?.imageSource) {
                    const src = c.webPartData.serverProcessedContent.imageSources.imageSource;
                    parts.push(`<div class="sp-news-image"><img src="${src}" alt="${item.Title}" /></div>`);
                }
            }

            return parts.join('\n');
        } catch {
            if (item.Description) return item.Description;
            return stripHtml(item.CanvasContent1 || '');
        }
    };

    const fullHtml = useMemo(() => (selectedNews ? buildFullHtml(selectedNews) : ''), [selectedNews]);

    return (
        <div className={styles.newsCollage}>
            <h2 className={styles.sectionTitle}>Latest News</h2>

            {news.length === 0 ? (
                <p className={styles.noItems}>No news found.</p>
            ) : (
                news.map((item: INewsItem) => (
                    <div
                        key={item.Id}
                        className={styles.newsCard}
                        onClick={() => setSelectedNews(item)}
                        role="button"
                        tabIndex={0}
                        onKeyDown={(e) => e.key === 'Enter' && setSelectedNews(item)}
                    >
                        {extractImageUrl(item) && (
                            <img
                                src={extractImageUrl(item)!}
                                alt={item.Title}
                                className={styles.newsImage}
                            />
                        )}

                        <div className={styles.newsContent}>
                            <h3 className={styles.newsTitle}>{item.Title}</h3>
                            <p className={styles.newsExcerpt}>
                                {item.Description
                                    ? item.Description.length > 140
                                        ? item.Description.substring(0, 140) + '…'
                                        : item.Description
                                    : item.CanvasContent1
                                        ? stripHtml(item.CanvasContent1).substring(0, 140) + '…'
                                        : ''}
                            </p>
                            <p className={styles.newsMeta}>
                                By {item.Author?.Title || 'Unknown'} on {new Date(item.Created).toLocaleDateString()}
                            </p>
                        </div>
                    </div>
                ))
            )}

            {selectedNews && (
                <div className={styles.modalOverlay} onClick={() => setSelectedNews(null)}>
                    <div className={styles.modalContent} onClick={(e) => e.stopPropagation()}>
                        <button className={styles.closeButton} onClick={() => setSelectedNews(null)} aria-label="Close">×</button>
                        <h2 className={styles.modalTitle}>{selectedNews.Title}</h2>
                        <p className={styles.newsMeta}>
                            By {selectedNews.Author?.Title || 'Unknown'} on {new Date(selectedNews.Created).toLocaleDateString()}
                        </p>
                        <div
                            className={styles.fullContent}
                            dangerouslySetInnerHTML={{ __html: fullHtml }}
                        />
                    </div>
                </div>
            )}
        </div>
    );
};

export default NewsList;