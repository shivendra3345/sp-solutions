import * as React from 'react';
import { useEffect, useState, useRef } from 'react';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './EventsCarousel.module.scss';

interface IEventItem {
    Title: string;
    Path?: string;
    SiteUrl?: string;
    Author?: string;
    EventDate?: string;
    EndDate?: string;
    Summary?: string;
    Banner?: string;
}

const RANDOM_BANNERS = [
    '/SiteAssets/images/random1.jpg',
    '/SiteAssets/images/random2.jpg',
    '/SiteAssets/images/random3.jpg'
];

const EventsCarousel: React.FC<{ context: any; excludedSitesCsv?: string; excludedListsCsv?: string; title?: string }> = ({ context, excludedSitesCsv = '', excludedListsCsv = '', title }) => {
    const [items, setItems] = useState<IEventItem[]>([]);
    const [index, setIndex] = useState(0);
    const mounted = useRef(true);

    useEffect(() => { return () => { mounted.current = false; }; }, []);

    useEffect(() => {
        const fetchEvents = async () => {
            // helper to extract a field value from different response shapes
            /*const extractField = (item: any, field: string) => {
                if (!item) return undefined;
                if (item[field] !== undefined && item[field] !== null) return item[field];
                const alt = field.replace(/^[A-Z]/, (m) => m.toUpperCase());
                if (item[alt] !== undefined) return item[alt];
                const ows = 'ows_' + field;
                if (item[ows] !== undefined) return item[ows];
                // check common array of Cells or Row/Cells shapes
                const cells = item.Cells || item.Row && item.Row.Cells || item.Row || item.Cells || item['Cell'] || item['Cells'];
                if (Array.isArray(cells)) {
                    const c = cells.find((x: any) => (x.Key || x.key || '').toString().toLowerCase() === field.toLowerCase() || (x.InternalName || x.internalName || '').toString().toLowerCase() === field.toLowerCase());
                    if (c) return c.Value || c.value;
                }
                // check FieldValues or FieldValuesAsText
                if (item.FieldValues && item.FieldValues[field] !== undefined) return item.FieldValues[field];
                if (item.FieldValuesAsText && item.FieldValuesAsText[field] !== undefined) return item.FieldValuesAsText[field];
                // sometimes RenderListDataAsStream returns attributes in a single property named like 'EventDate'
                if (item['EventDate']) return item['EventDate'];
                return undefined;
            };*/

            try {
                // Enumerate webs in the site collection and then calendar lists (BaseTemplate 106)
                // from each web. This avoids the Search REST API and works via List REST endpoints.
                const siteRoot = context.pageContext.site && context.pageContext.site.absoluteUrl ? context.pageContext.site.absoluteUrl : context.pageContext.web.absoluteUrl;

                // BFS to collect all web urls (include root)
                const webUrls: string[] = [siteRoot];
                for (let i = 0; i < webUrls.length; i++) {
                    try {
                        const wurl = `${webUrls[i]}/_api/web/webs?$select=Url`;
                        const r = await context.spHttpClient.get(wurl, SPHttpClient.configurations.v1);
                        if (!r || !r.ok) continue;
                        const j = await r.json();
                        const subs = j.value || [];
                        subs.forEach((s: any) => {
                            if (s && s.Url) webUrls.push(s.Url);
                        });
                    } catch (err) {
                        // ignore web enumeration errors for individual webs
                        console.debug('EventsCarousel: failed to enumerate subwebs for', webUrls[i], err);
                    }
                }
                console.debug('EventsCarousel: discovered webUrls count', webUrls.length, webUrls.slice(0, 20));

                const events: IEventItem[] = [];
                const seen = new Set<string>();
                const today = new Date();
                today.setHours(0, 0, 0, 0);
                for (const webUrl of webUrls) {
                    try {
                        // find candidate lists (non-hidden). We'll accept lists that are BaseTemplate 106 (calendar)
                        // or lists that contain an EventDate field (custom event lists).
                        const listsUrl = `${webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate`;
                        const lr = await context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1);
                        if (!lr || !lr.ok) continue;
                        const lj = await lr.json();
                        const allLists = lj.value || [];
                        const lists: any[] = [];
                        // prepare excluded lists set from prop
                        const excludedLists = (excludedListsCsv || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
                        for (const l of allLists) {
                            try {
                                const id = (l.Id || l.ID || l.id || '').toString();
                                const title = (l.Title || l.Title || '').toString();
                                const base = l.BaseTemplate || l.BaseTemplate || 0;
                                // skip excluded lists by id or title
                                if (excludedLists.some(e => e === id.toLowerCase() || e === title.toLowerCase())) continue;
                                if (base === 106) {
                                    lists.push(l);
                                    continue;
                                }
                                // otherwise probe fields for EventDate presence
                                try {
                                    const fieldsUrl = `${webUrl}/_api/web/lists(guid'${id}')/fields?$select=InternalName,Title&$filter=InternalName eq 'EventDate' or Title eq 'EventDate'`;
                                    const fr = await context.spHttpClient.get(fieldsUrl, SPHttpClient.configurations.v1);
                                    if (fr && fr.ok) {
                                        const fj = await fr.json();
                                        const fvals = fj.value || [];
                                        if (Array.isArray(fvals) && fvals.length > 0) {
                                            lists.push(l);
                                        }
                                    }
                                } catch (ferr) {
                                    // ignore field-probe errors
                                }
                            } catch (inner) {
                                // ignore per-list errors
                            }
                        }
                        console.debug('EventsCarousel: found lists for web', webUrl, lists.length, lists.map((l: any) => l.Title));
                        // fetch a form digest for this web (required for POST GetItems)
                        let formDigest: string | undefined = undefined;
                        for (const list of lists) {
                            try {
                                // Use CAML GetItems with ExpandRecurrence to expand recurring events
                                const rangeDays = 90; // look ahead window
                                const start = new Date(); start.setHours(0, 0, 0, 0);
                                const end = new Date(start.getTime()); end.setDate(end.getDate() + rangeDays);
                                const startIso = start.toISOString();
                                const endIso = end.toISOString();
                                const caml = `
                                                                        <View>
                                                                            <Query>
                                                                                <Where>
                                                                                    <And>
                                                                                        <Geq><FieldRef Name='EventDate'/><Value Type='DateTime'>${startIso}</Value></Geq>
                                                                                        <Leq><FieldRef Name='EventDate'/><Value Type='DateTime'>${endIso}</Value></Leq>
                                                                                    </And>
                                                                                </Where>
                                                                                <OrderBy><FieldRef Name='EventDate' Ascending='TRUE'/></OrderBy>
                                                                            </Query>
                                                                            <ViewFields>
                                                                                <FieldRef Name='Title'/>
                                                                                <FieldRef Name='EventDate'/>
                                                                                <FieldRef Name='EndDate'/>
                                                                                <FieldRef Name='Author'/>
                                                                                <FieldRef Name='ID'/>
                                                                                <FieldRef Name='FileRef'/>
                                                                            </ViewFields>
                                                                            <RowLimit>500</RowLimit>
                                                                            <QueryOptions>
                                                                                <ExpandRecurrence>TRUE</ExpandRecurrence>
                                                                                <CalendarDate>${startIso}</CalendarDate>
                                                                            </QueryOptions>
                                                                        </View>`;

                                // Try RenderListDataAsStream which is more consistent across tenants
                                const renderUrl = `${webUrl}/_api/web/lists(guid'${list.Id}')/RenderListDataAsStream`;
                                const renderBody = JSON.stringify({ parameters: { ViewXml: caml } });

                                // Ensure we have a form digest for POST requests
                                if (!formDigest) {
                                    try {
                                        const ctxRes = await context.spHttpClient.post(`${webUrl}/_api/contextinfo`, SPHttpClient.configurations.v1, {
                                            headers: { 'Accept': 'application/json;odata=verbose' },
                                            body: ''
                                        });
                                        if (ctxRes && ctxRes.ok) {
                                            const ctxJson = await ctxRes.json();
                                            formDigest = (ctxJson && ctxJson.d && ctxJson.d.GetContextWebInformation && ctxJson.d.GetContextWebInformation.FormDigestValue) || ctxJson.FormDigestValue || undefined;
                                        }
                                    } catch (dex) {
                                        console.debug('EventsCarousel: failed to get form digest for', webUrl, dex);
                                    }
                                }

                                let itemsArray: any[] = [];
                                try {
                                    const rr = await context.spHttpClient.post(renderUrl, SPHttpClient.configurations.v1, {
                                        headers: {
                                            'Accept': 'application/json;odata=verbose',
                                            'Content-Type': 'application/json;odata=verbose',
                                            ...(formDigest ? { 'X-RequestDigest': formDigest } : {})
                                        },
                                        body: renderBody
                                    });
                                    if (rr && rr.ok) {
                                        const rj = await rr.json();
                                        console.debug('EventsCarousel: RenderListDataAsStream response keys', Object.keys(rj || {}));
                                        // Try common response shapes
                                        if (rj && rj.Row) itemsArray = Array.isArray(rj.Row) ? rj.Row : (rj.Row.results || []);
                                        else if (rj && rj.Rows) itemsArray = Array.isArray(rj.Rows) ? rj.Rows : (rj.Rows.results || []);
                                        else if (rj && rj.value) itemsArray = rj.value;
                                        else if (rj && rj.d && rj.d.results) itemsArray = rj.d.results;
                                        else itemsArray = [];
                                        console.debug('EventsCarousel: RenderListDataAsStream returned', itemsArray.length, 'items for list', list && list.Title);
                                        if (itemsArray.length > 0) {
                                            // log a sample item to inspect structure for debugging
                                            try { console.debug('EventsCarousel: sample item', JSON.stringify(itemsArray[0], null, 2)); } catch (e) { console.debug('EventsCarousel: sample item (raw)', itemsArray[0]); }
                                        }
                                    } else {
                                        const txt = await (rr && rr.text ? rr.text() : Promise.resolve(''));
                                        console.debug('EventsCarousel: RenderListDataAsStream failed', rr && rr.status, txt);
                                    }
                                } catch (errR) {
                                    console.debug('EventsCarousel: RenderListDataAsStream threw', errR);
                                }

                                // If RenderListDataAsStream returned no items, fallback to OData GET filtered by EventDate
                                if ((!itemsArray || itemsArray.length === 0)) {
                                    try {
                                        const isoNoZ = start.toISOString().split('.')[0];
                                        const fallbackUrl = `${webUrl}/_api/web/lists(guid'${list.Id}')/items?$select=Title,EventDate,EndDate,Author/Title,ID,FileRef&$expand=Author&$filter=EventDate ge datetime'${isoNoZ}'&$top=50&$orderby=EventDate asc`;
                                        console.debug('EventsCarousel: Render returned 0, trying fallback OData GET', fallbackUrl);
                                        const fr = await context.spHttpClient.get(fallbackUrl, SPHttpClient.configurations.v1);
                                        if (fr && fr.ok) {
                                            const fj = await fr.json();
                                            const fitems = fj && (fj.value || fj.d && fj.d.results) || [];
                                            console.debug('EventsCarousel: fallback returned', (fitems || []).length, 'items for list', list && list.Title);
                                            if (Array.isArray(fitems) && fitems.length) {
                                                itemsArray = itemsArray.concat(fitems);
                                            }
                                        }
                                    } catch (ferr) {
                                        console.debug('EventsCarousel: fallback OData failed for list', list && list.Title, ferr);
                                    }
                                }

                                for (const it of itemsArray) {
                                    try {
                                        // only include today or future events
                                        const evDate = it.EventDate ? new Date(it.EventDate) : null;
                                        if (!evDate) continue;
                                        if (evDate < today) continue;

                                        const key = `${(it.Title || '').toString().trim()}|${evDate.toISOString()}|${webUrl}`;
                                        if (seen.has(key)) continue;
                                        seen.add(key);

                                        const origin = (typeof window !== 'undefined' && window.location && window.location.origin) ? window.location.origin : (new URL(context.pageContext.web.absoluteUrl)).origin;
                                        const fileRef = it.FileRef || '';
                                        const fullPathCandidate = fileRef && fileRef.indexOf('/') === 0 ? `${origin}${fileRef}` : (fileRef || '');
                                        const displayUrl = `${webUrl}/_layouts/15/listform.aspx?PageType=4&ListId=${list.Id}&ID=${it.ID}`;
                                        const finalPath = fullPathCandidate && fullPathCandidate.length > 0 ? fullPathCandidate : displayUrl;

                                        const evt: IEventItem = {
                                            Title: it.Title || 'Untitled',
                                            Path: finalPath,
                                            SiteUrl: webUrl,
                                            Author: it.Author && it.Author.Title ? it.Author.Title : undefined,
                                            EventDate: it.EventDate,
                                            EndDate: it.EndDate,
                                            Summary: undefined,
                                            Banner: undefined
                                        };

                                        // try to fetch first image attachment for the item (if any)
                                        try {
                                            const attachUrl = `${webUrl}/_api/web/lists(guid'${list.Id}')/items(${it.ID})/AttachmentFiles`;
                                            const ar = await context.spHttpClient.get(attachUrl, SPHttpClient.configurations.v1);
                                            if (ar && ar.ok) {
                                                const aj = await ar.json();
                                                const af = (aj.value || []).find((a: any) => /\.(jpe?g|png|gif)$/i.test(a.FileName || ''));
                                                if (af && af.ServerRelativeUrl) {
                                                    evt.Banner = `${origin}${af.ServerRelativeUrl}`;
                                                }
                                            }
                                        } catch (aerr) {
                                            /* ignore attachment errors */
                                        }

                                        events.push(evt);
                                    } catch (err) {
                                        console.debug('EventsCarousel: failed processing item', it && it.ID, err);
                                    }
                                }
                            } catch (err) {
                                console.debug('EventsCarousel: failed to fetch items for list', list && list.Title, err);
                            }
                        }
                    } catch (err) {
                        console.debug('EventsCarousel: failed to list calendars in', webUrl, err);
                    }
                }

                // apply excludedSites filter
                const excluded = (excludedSitesCsv || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
                const filtered = events.filter(it => {
                    if (!it.SiteUrl) return true;
                    const site = it.SiteUrl.toString().toLowerCase();
                    return !excluded.some(e => site.indexOf(e) !== -1 || e.indexOf(site) !== -1);
                });

                if (mounted.current) setItems(filtered.slice(0, 30));
            } catch (e) {
                console.error('EventsCarousel: REST enumeration failed', e);
            }
        };

        fetchEvents();
    }, [context, excludedSitesCsv]);

    // determine visible slides: if fewer than 3 items, show them as-is (no duplication).
    let visible: IEventItem[] = [];
    if (!items || items.length === 0) visible = [];
    else if (items.length < 3) visible = items;
    else {
        visible = [];
        for (let i = 0; i < 3; i++) {
            visible.push(items[(index + i) % items.length]);
        }
    }

    const prev = () => {
        if (!items || items.length === 0) return;
        setIndex((idx) => (idx - 1 + items.length) % items.length);
    };
    const next = () => {
        if (!items || items.length === 0) return;
        setIndex((idx) => (idx + 1) % items.length);
    };

    const randomBanner = (i: number) => {
        const pick = Math.abs(i) % RANDOM_BANNERS.length;
        return RANDOM_BANNERS[pick];
    };

    return (
        <div>
            <h3 className={styles.title}>{title || 'Events'}</h3>
            <div className={styles.carouselWrap}>
                <button className={styles.arrowLeft} aria-label="Previous" onClick={prev}>‹</button>
                <div className={styles.slides}>
                    {visible.length === 0 && (<div className={styles.empty}>No upcoming events</div>)}
                    {visible.map((it, i) => (
                        <a key={`${it.Path}-${i}`} className={styles.slide} href={it.Path || it.SiteUrl} target="_blank" rel="noreferrer">
                            <div className={styles.banner} style={{ backgroundImage: `url('${it.Banner || randomBanner(i)}')` }} />
                            <div className={styles.meta}>
                                <h4 className={styles.title}>{it.Title}</h4>
                                <div className={styles.info}>By {it.Author || 'Unknown'}</div>
                                <div className={styles.info}>{it.EventDate ? (new Date(it.EventDate)).toLocaleString() : ''}{it.EndDate ? ` - ${new Date(it.EndDate).toLocaleString()}` : ''}</div>
                                {it.Summary && <p className={styles.summary}>{it.Summary.replace(/(<([^>]+)>)/gi, '').slice(0, 220)}</p>}
                            </div>
                        </a>
                    ))}
                </div>
                <button className={styles.arrowRight} aria-label="Next" onClick={next}>›</button>
            </div>
        </div>
    );
};

export default EventsCarousel;
