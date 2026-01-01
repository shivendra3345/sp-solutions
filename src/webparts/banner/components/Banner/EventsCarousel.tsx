import * as React from 'react';
import { useEffect, useState, useRef } from 'react';
import { Dialog, DialogType, DialogFooter } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
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
    ListId?: string;
    ItemId?: number;
}

const RANDOM_BANNERS = [
    'random1.jpg',
    'random2.jpg'

];

const EventsCarousel: React.FC<{ context: any; excludedSitesCsv?: string; excludedListsCsv?: string; selectedListIdsCsv?: string; descriptionFieldNamesCsv?: string; title?: string }> = ({ context, excludedSitesCsv = '', excludedListsCsv = '', selectedListIdsCsv = '', descriptionFieldNamesCsv = '', title }) => {
    const [items, setItems] = useState<IEventItem[]>([]);
    const [index, setIndex] = useState(0);
    const mounted = useRef(true);
    const [selectedEvent, setSelectedEvent] = useState<IEventItem | null>(null);
    const [isModalOpen, setIsModalOpen] = useState(false);
    const MODAL_HOST_ID = 'sp-solutions-modal-host';

    // Ensure a top-level modal host exists so Fluent UI Dialog/Layer renders outside any transformed/overflowed ancestors
    React.useEffect(() => {
        if (typeof document === 'undefined') return;
        let host = document.getElementById(MODAL_HOST_ID);
        if (!host) {
            host = document.createElement('div');
            host.id = MODAL_HOST_ID;
            document.body.appendChild(host);
        }
        return () => {
            // keep host — don't remove it to avoid interfering with other components
        };
    }, []);

    useEffect(() => { return () => { mounted.current = false; }; }, []);

    useEffect(() => {
        const fetchEvents = async () => {
            const events: IEventItem[] = [];
            try {
                const siteRoot = context.pageContext.site && context.pageContext.site.absoluteUrl ? context.pageContext.site.absoluteUrl : context.pageContext.web.absoluteUrl;
                // const origin = (typeof window !== 'undefined' && window.location && window.location.origin)
                //     ? window.location.origin
                //     : (new URL(context.pageContext.web.absoluteUrl)).origin;
                const seen = new Set<string>();
                const today = new Date(); today.setHours(0, 0, 0, 0);

                // simple helper to read date-like values from various shapes
                const getFieldValue = (item: any, fieldName: string) => {
                    if (!item) return undefined;
                    if (item[fieldName] !== undefined && item[fieldName] !== null) return item[fieldName];
                    if (item.FieldValuesAsText && item.FieldValuesAsText[fieldName] !== undefined) return item.FieldValuesAsText[fieldName];
                    const ows = 'ows_' + fieldName;
                    if (item[ows] !== undefined) return item[ows];
                    const cells = item.Cells || item.Row && item.Row.Cells || item.Cell || item.Cells;
                    if (Array.isArray(cells)) {
                        const c = cells.find((x: any) => (x.InternalName || x.Name || x.Key || '').toString().toLowerCase().indexOf(fieldName.toLowerCase()) !== -1);
                        if (c) return c.Value || c.value;
                    }
                    for (const k of Object.keys(item)) {
                        if (k.toLowerCase().indexOf('date') !== -1 && item[k]) return item[k];
                    }
                    return undefined;
                };

                // build target lists: either from selectedListIdsCsv or by discovery
                const targetLists: any[] = [];
                if (selectedListIdsCsv && selectedListIdsCsv.trim()) {
                    const entries = (selectedListIdsCsv || '').split(',').map(s => s.trim()).filter(Boolean);
                    for (const e of entries) {
                        const parts = e.split('::');
                        if (parts.length !== 2) continue;
                        const webUrl = decodeURIComponent(parts[0]);
                        const listId = parts[1];
                        targetLists.push({ Id: listId, __webUrl: webUrl });
                    }
                } else {
                    const webUrls: string[] = [siteRoot];
                    for (let i = 0; i < webUrls.length; i++) {
                        try {
                            const wurl = `${webUrls[i]}/_api/web/webs?$select=Url`;
                            const r = await context.spHttpClient.get(wurl, SPHttpClient.configurations.v1);
                            if (!r || !r.ok) continue;
                            const j = await r.json();
                            const subs = j.value || [];
                            subs.forEach((s: any) => { if (s && s.Url) webUrls.push(s.Url); });
                        } catch (err) { /* ignore */ }
                    }
                    for (const webUrl of webUrls) {
                        try {
                            const listsUrl = `${webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate`;
                            const lr = await context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1);
                            if (!lr || !lr.ok) continue;
                            const lj = await lr.json();
                            const allLists = lj.value || [];
                            const excludedLists = (excludedListsCsv || '').split(',').map(s => s.trim().toLowerCase()).filter(Boolean);
                            for (const l of allLists) {
                                const id = (l.Id || l.ID || l.id || '').toString();
                                const title = (l.Title || '').toString();
                                const base = l.BaseTemplate || 0;
                                if (excludedLists.some(e => e === id.toLowerCase() || e === title.toLowerCase())) continue;
                                if (base === 106 || title.toLowerCase().indexOf('event') !== -1 || title.toLowerCase().indexOf('calendar') !== -1) {
                                    l.__webUrl = webUrl;
                                    targetLists.push(l);
                                }
                            }
                        } catch (err) { /* ignore */ }
                    }
                }

                // probe lists for date fields (Start/End/EventDate) when possible
                const probeDateFields = async (list: any) => {
                    try {
                        const webUrl = list.__webUrl || siteRoot;
                        const listId = list.Id || list.ID;
                        const fUrl = `${webUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false and TypeAsString eq 'DateTime'&$select=InternalName,Title`;
                        const fr = await context.spHttpClient.get(fUrl, SPHttpClient.configurations.v1);
                        if (!fr || !fr.ok) return;
                        const fj = await fr.json();
                        const fitems = fj && (fj.value || fj.d && fj.d.results) || [];
                        list.__dateFields = fitems.map((f: any) => ({ InternalName: f.InternalName, Title: f.Title }));
                        console.debug('EventsCarousel: probeDateFields', { listId: listId, webUrl: webUrl, dateFields: list.__dateFields });
                    } catch (e) {
                        // ignore probe errors
                    }
                };

                // probe lists for multiline/Note fields (likely to contain descriptions)
                const probeRichTextFields = async (list: any) => {
                    try {
                        const webUrl = list.__webUrl || siteRoot;
                        const listId = list.Id || list.ID;
                        const fUrl = `${webUrl}/_api/web/lists(guid'${listId}')/fields?$filter=Hidden eq false and TypeAsString eq 'Note'&$select=InternalName,Title`;
                        const fr = await context.spHttpClient.get(fUrl, SPHttpClient.configurations.v1);
                        if (!fr || !fr.ok) return;
                        const fj = await fr.json();
                        const fitems = fj && (fj.value || fj.d && fj.d.results) || [];
                        list.__richTextFields = fitems.map((f: any) => ({ InternalName: f.InternalName, Title: f.Title }));
                        console.debug('EventsCarousel: probeRichTextFields', { listId: listId, webUrl: webUrl, richTextFields: list.__richTextFields });
                    } catch (e) {
                        // ignore
                    }
                };

                for (const list of targetLists) {
                    try {
                        const webUrl = list.__webUrl || siteRoot;
                        const listId = list.Id || list.ID;
                        // ensure we have date field metadata for selected lists
                        if (!list.__dateFields) await probeDateFields(list);
                        // if probing returned no DateTime fields, skip this list to avoid querying non-existent fields
                        if (!list.__dateFields || list.__dateFields.length === 0) {
                            console.debug('EventsCarousel: skipping list (no DateTime fields)', { listId: listId, webUrl });
                            continue;
                        }
                        const isoNoZ = (new Date()).toISOString().split('.')[0];
                        // attempt to pick common start/end fields; fall back to EventDate for start
                        const preferredStart = ['EventDate', 'StartDate', 'Start', 'BeginDate', 'Published_x0020_Date', 'Event_x0020_Date', 'Date'];
                        const preferredEnd = ['EndDate', 'End'];
                        let startField = 'EventDate';
                        let endField: string | undefined = undefined;
                        if (list.__dateFields && list.__dateFields.length) {
                            const foundStart = list.__dateFields.find((d: any) => preferredStart.map(p => p.toLowerCase()).indexOf((d.InternalName || '').toLowerCase()) !== -1 || preferredStart.map(p => p.toLowerCase()).indexOf((d.Title || '').toLowerCase()) !== -1);
                            const foundEnd = list.__dateFields.find((d: any) => preferredEnd.map(p => p.toLowerCase()).indexOf((d.InternalName || '').toLowerCase()) !== -1 || preferredEnd.map(p => p.toLowerCase()).indexOf((d.Title || '').toLowerCase()) !== -1);
                            startField = (foundStart && foundStart.InternalName) || (list.__dateFields[0] && list.__dateFields[0].InternalName) || startField;
                            endField = (foundEnd && foundEnd.InternalName) || undefined;
                            console.debug('EventsCarousel: chosen date fields', { listId: listId, startField, endField });
                        }
                        const selectFields = ['Title', startField, 'Author/Title', 'ID', 'FileRef', 'AttachmentFiles/ServerRelativeUrl'];
                        // include common candidate fields that may contain the event description/html body
                        const preferredDesc = ['Description', 'Body', 'EventDescription', 'Comments', 'Summary', 'CanvasContent1', 'Event_x0020_Description'];
                        // ensure we probed rich text fields for this list and only add candidates that actually exist
                        if (!list.__richTextFields) await probeRichTextFields(list);
                        const availableDescInternal = (list.__richTextFields || []).map((f: any) => (f.InternalName || '').toString().toLowerCase());
                        const availableDescTitles = (list.__richTextFields || []).map((f: any) => (f.Title || '').toString().toLowerCase());
                        for (const d of preferredDesc) {
                            const dLower = d.toLowerCase();
                            // if the list has a field whose internal name or title matches our candidate, include it
                            if (availableDescInternal.indexOf(dLower) !== -1 || availableDescTitles.indexOf(dLower) !== -1) {
                                // find the actual internal name from the list
                                const found = (list.__richTextFields || []).find((f: any) => ((f.InternalName || '').toString().toLowerCase() === dLower) || ((f.Title || '').toString().toLowerCase() === dLower));
                                if (found && selectFields.indexOf(found.InternalName) === -1) selectFields.push(found.InternalName);
                            }
                        }
                        // always include FieldValuesAsText so we can read text representations safely
                        if (selectFields.indexOf('FieldValuesAsText') === -1) selectFields.push('FieldValuesAsText');
                        if (endField) selectFields.push(endField);
                        const expandFields = ['Author', 'AttachmentFiles'];
                        const fallbackUrl = `${webUrl}/_api/web/lists(guid'${listId}')/items?$select=${selectFields.join(',')}&$expand=${expandFields.join(',')}&$filter=${startField} ge datetime'${isoNoZ}'&$top=200&$orderby=${startField} asc`;
                        const fr = await context.spHttpClient.get(fallbackUrl, SPHttpClient.configurations.v1);
                        if (!fr || !fr.ok) continue;
                        const fj = await fr.json();
                        const fitems = fj && (fj.value || fj.d && fj.d.results) || [];
                        let sampleLogged = false;
                        for (const it of fitems) {
                            try {
                                if (!sampleLogged) {
                                    console.debug('EventsCarousel: sample item for list', { listId, sample: it });
                                    sampleLogged = true;
                                }
                                const evRaw = getFieldValue(it, startField) || it[startField] || (it.EventDate);
                                const evDate = evRaw ? new Date(evRaw) : null;
                                const evEndRaw = endField ? (getFieldValue(it, endField) || it[endField]) : undefined;
                                // const evEndDate = evEndRaw ? new Date(evEndRaw) : undefined;
                                if (!evDate) continue;
                                if (evDate < today) continue;
                                const key = `${(it.Title || '').toString().trim()}|${evDate.toISOString()}|${webUrl}`;
                                if (seen.has(key)) continue;
                                seen.add(key);
                                const displayUrl = `${webUrl}/_layouts/15/listform.aspx?PageType=4&ListId=${listId}&ID=${it.ID || it.Id}`;
                                let bannerUrl: string | undefined = undefined;
                                try {
                                    if (it.AttachmentFiles && it.AttachmentFiles.length) {
                                        const af = it.AttachmentFiles[0];
                                        bannerUrl = af.ServerRelativeUrl || af.FileName && `${webUrl.replace(/\/$/, '')}/Lists/${encodeURIComponent(list.Title)}/Attachments/${it.ID || it.Id}/${af.FileName}`;
                                    } else if (it.FileRef && typeof it.FileRef === 'string' && /\.(jpg|jpeg|png|gif)$/i.test(it.FileRef)) {
                                        bannerUrl = it.FileRef;
                                    }
                                } catch (e) { /* ignore banner resolution errors */ }

                                // pick the first available description field value. Prefer actual list fields discovered above.
                                let descRaw: string | undefined = undefined;
                                const discoveredDescFields = (list.__richTextFields || []).map((f: any) => f.InternalName).filter(Boolean);
                                const userProvided = (descriptionFieldNamesCsv || '').split(',').map(s => s.trim()).filter(Boolean);
                                const descCandidates = [...userProvided, ...discoveredDescFields, 'Description', 'Body', 'EventDescription', 'Comments', 'Summary', 'CanvasContent1', 'Event_x0020_Description'];
                                for (const d of descCandidates) {
                                    try {
                                        const v = getFieldValue(it, d) || (it.FieldValuesAsText && it.FieldValuesAsText[d]) || it[d];
                                        if (v) { descRaw = v; break; }
                                    } catch (e) { /* ignore per-field */ }
                                }

                                const evt: IEventItem = {
                                    Title: it.Title || 'Untitled',
                                    Path: displayUrl,
                                    SiteUrl: webUrl,
                                    Author: (it.Author && it.Author.Title) || (it.FieldValuesAsText && it.FieldValuesAsText.Author) || undefined,
                                    EventDate: evRaw,
                                    EndDate: evEndRaw,
                                    Banner: bannerUrl,
                                    Summary: descRaw,
                                    ListId: listId,
                                    ItemId: it.ID || it.Id || undefined
                                };
                                events.push(evt);
                            } catch (inner) { /* ignore per-item errors */ }
                        }
                    } catch (err) { console.debug('EventsCarousel: failed to fetch items for list', list && (list.Title || list.Id), err); }
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
        return () => { mounted.current = false; };
    }, [context, excludedSitesCsv, selectedListIdsCsv]);

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

    const showArrows = !!items && items.length > 0 && !!visible && visible.length > 1;

    const prev = () => {
        if (!items || items.length === 0) return;
        setIndex((idx) => (idx - 1 + items.length) % items.length);
    };
    const next = () => {
        if (!items || items.length === 0) return;
        setIndex((idx) => (idx + 1) % items.length);
    };

    // Deterministic banner selection per-event so images are stable per event
    const hashString = (s: string) => {
        let h = 0;
        for (let i = 0; i < s.length; i++) {
            h = ((h << 5) - h) + s.charCodeAt(i);
            h |= 0; // convert to 32bit int
        }
        return Math.abs(h);
    };

    const bannerForEvent = (evt: IEventItem) => {
        if (evt.Banner && typeof evt.Banner === 'string') {
            // If Banner is already an absolute/relative url, use it as-is
            return evt.Banner;
        }
        // build a deterministic pick based on event identity
        const key = `${evt.SiteUrl || ''}::${evt.ListId || ''}::${evt.ItemId || ''}::${evt.Title || ''}`;
        const pick = hashString(key) % RANDOM_BANNERS.length;
        const name = RANDOM_BANNERS[pick];
        const base = evt.SiteUrl ? evt.SiteUrl.replace(/\/$/, '') : (typeof window !== 'undefined' ? window.location.origin : '');
        return `${base}/SiteAssets/images/${name}`;
    };

    const formatDateFriendly = (d?: string | null) => {
        if (!d) return '';
        try {
            const date = new Date(d);
            return date.toLocaleString(undefined, { year: 'numeric', month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' });
        } catch (e) { return d; }
    };

    const stripHtml = (s?: string) => {
        if (!s) return '';
        try { return s.replace(/(<([^>]+)>)/gi, ''); } catch (e) { return s; }
    };
    try { console.debug('EventsCarousel: counts', { itemsCount: items && items.length, visibleCount: visible && visible.length }); } catch (e) { }

    return (
        <div>
            <h3 className={styles.title}>{title || 'Events'}</h3>
            {(!visible || visible.length === 0) ? (
                <div className={styles.carouselWrap}>
                    <div className={styles.slides}>
                        <div className={styles.empty}>No upcoming events</div>
                    </div>
                </div>
            ) : (
                <div className={styles.carouselWrap}>
                    {showArrows && <button className={styles.arrowLeft} aria-label="Previous" onClick={prev}>‹</button>}
                    <div className={styles.slides}>
                        {visible.map((it, i) => (
                            <div key={`${it.Path || it.Title}-${i}`} className={styles.slide} role="button" tabIndex={0}
                                onClick={(e) => { e.preventDefault(); setSelectedEvent(it); setIsModalOpen(true); }}
                                onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); setSelectedEvent(it); setIsModalOpen(true); } }}
                            >
                                <div className={styles.banner} style={{ backgroundImage: `url('${bannerForEvent(it)}')` }} />
                                <div className={styles.meta} style={{ textAlign: 'left' }}>
                                    <h4 className={styles.title} style={{ textAlign: 'left' }}>{it.Title}</h4>
                                    <div className={styles.info} style={{ textAlign: 'left' }}>
                                        {it.EventDate ? formatDateFriendly(it.EventDate) : ''}{it.EndDate ? ` - ${formatDateFriendly(it.EndDate)}` : ''}
                                    </div>
                                    <div className={styles.info} style={{ textAlign: 'left' }}>By {it.Author || 'Unknown'}</div>
                                    {it.Summary && (
                                        <p className={styles.summary} style={{
                                            display: '-webkit-box',
                                            WebkitLineClamp: 2,
                                            WebkitBoxOrient: 'vertical',
                                            overflow: 'hidden',
                                            textOverflow: 'ellipsis',
                                            marginTop: 8
                                        }}>{stripHtml(it.Summary)}</p>
                                    )}
                                </div>
                            </div>
                        ))}
                    </div>
                    {showArrows && <button className={styles.arrowRight} aria-label="Next" onClick={next}>›</button>}
                </div>
            )}
            {/* Event details dialog */}
            <Dialog
                hidden={!isModalOpen}
                onDismiss={() => { setIsModalOpen(false); setSelectedEvent(null); }}
                dialogContentProps={{
                    type: DialogType.largeHeader,
                    title: selectedEvent ? selectedEvent.Title : 'Event Details',
                    subText: selectedEvent && (selectedEvent.EventDate || selectedEvent.EndDate) ? `${selectedEvent.EventDate ? formatDateFriendly(selectedEvent.EventDate) : ''}${selectedEvent.EndDate ? ` - ${formatDateFriendly(selectedEvent.EndDate)}` : ''}` : undefined
                }}
                modalProps={{ isBlocking: false, layerProps: { hostId: MODAL_HOST_ID } }}
            >
                <div>
                    {selectedEvent && (
                        <div>
                            {selectedEvent.Banner && <img src={selectedEvent.Banner} alt={selectedEvent.Title} className={styles.modalImage} />}
                            <div className={styles.modalBody}>
                                <div style={{ marginBottom: 8 }}>
                                    <div style={{ marginTop: 6 }}><strong>By:</strong> {selectedEvent.Author || 'Unknown'}</div>
                                    {selectedEvent.SiteUrl && <div style={{ marginTop: 6 }}><strong>Site:</strong> <a href={selectedEvent.SiteUrl} target="_blank" rel="noreferrer">{selectedEvent.SiteUrl}</a></div>}
                                </div>
                                <div>
                                    {selectedEvent.Summary && <div dangerouslySetInnerHTML={{ __html: selectedEvent.Summary }} />}
                                </div>
                            </div>
                        </div>
                    )}
                </div>
                <DialogFooter>
                    <PrimaryButton onClick={() => {
                        if (!selectedEvent) return;
                        // Prefer constructing a display form URL from SiteUrl + ListId + ItemId
                        try {
                            let openUrl = selectedEvent.Path;
                            if (selectedEvent.SiteUrl && selectedEvent.ListId && selectedEvent.ItemId) {
                                openUrl = `${selectedEvent.SiteUrl}/_layouts/15/listform.aspx?PageType=4&ListId=${selectedEvent.ListId}&ID=${selectedEvent.ItemId}`;
                            }
                            console.debug('EventsCarousel: opening event URL', openUrl, selectedEvent);
                            window.open(openUrl, '_blank');
                        } catch (e) {
                            console.error('EventsCarousel: failed to open event url', e, selectedEvent);
                        }
                    }} text="Open event" />
                    <DefaultButton onClick={() => { setIsModalOpen(false); setSelectedEvent(null); }} text="Close" />
                </DialogFooter>
            </Dialog>
        </div>
    );
};

export default EventsCarousel;
