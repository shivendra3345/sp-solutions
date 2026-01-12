import * as React from 'react';
import { useEffect, useState } from 'react';
//import * as ReactDom from 'react-dom';
import { SPHttpClient } from '@microsoft/sp-http';

interface IListEntry {
    id: string;
    title: string;
    webUrl: string;
}

interface IProps {
    context: any;
    currentSelectedCsv?: string;
    onSave: (csv: string, titlesCsv: string) => void;
    onDismiss: () => void;
}

const CalendarPicker: React.FC<IProps> = ({ context, currentSelectedCsv = '', onSave, onDismiss }) => {
    const [items, setItems] = useState<IListEntry[]>([]);
    const [selected, setSelected] = useState<Record<string, boolean>>({});
    const [loading, setLoading] = useState(false);
    const [error, setError] = useState<string | null>(null);

    useEffect(() => {
        const initial = (currentSelectedCsv || '').split(',').map(s => s.trim()).filter(Boolean);
        const sel: Record<string, boolean> = {};
        initial.forEach(i => sel[i] = true);
        setSelected(sel);
    }, [currentSelectedCsv]);

    useEffect(() => {
        let cancelled = false;
        const discover = async () => {
            setLoading(true);
            setError(null);
            try {
                console.debug('CalendarPicker: starting discovery');
                const siteRoot = context.pageContext.site && context.pageContext.site.absoluteUrl ? context.pageContext.site.absoluteUrl : context.pageContext.web.absoluteUrl;
                console.debug('CalendarPicker: siteRoot', siteRoot);
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

                const discovered: IListEntry[] = [];
                for (const webUrl of webUrls) {
                    try {
                        const listsUrl = `${webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate`;
                        const lr = await context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1);
                        if (!lr || !lr.ok) continue;
                        const lj = await lr.json();
                        const allLists = lj.value || [];
                        for (const l of allLists) {
                            try {
                                const id = (l.Id || l.ID || l.id || '').toString();
                                const title = (l.Title || '').toString();
                                const base = l.BaseTemplate || 0;
                                if (base === 106 || title.toLowerCase().indexOf('event') !== -1 || title.toLowerCase().indexOf('calendar') !== -1) {
                                    discovered.push({ id, title, webUrl });
                                }
                            } catch (inner) { }
                        }
                    } catch (e) { /* ignore */ }
                }

                if (!cancelled) {
                    setItems(discovered);
                    console.debug('CalendarPicker: discovered lists count', discovered.length);
                }
            } catch (e) {
                console.error('CalendarPicker: discover failed', e);
                if (!cancelled) setError((e && e.message) || String(e));
            } finally {
                if (!cancelled) setLoading(false);
            }
        };
        discover();
        return () => { cancelled = true; };
    }, [context]);

    const toggle = (entryKey: string) => {
        setSelected(prev => ({ ...prev, [entryKey]: !prev[entryKey] }));
    };

    const doSave = () => {
        const selectedKeys = Object.keys(selected).filter(k => selected[k]);
        const titles: string[] = [];
        for (const k of selectedKeys) {
            const found = items.find(it => `${encodeURIComponent(it.webUrl)}::${it.id}` === k);
            if (found) titles.push(`${found.title} (${found.webUrl})`);
            else titles.push(k);
        }
        onSave(selectedKeys.join(','), titles.join(', '));
    };

    return (
        <div style={{ position: 'fixed', zIndex: 2147483647, top: '10%', left: '50%', transform: 'translateX(-50%)', width: '720px', maxWidth: '94%', background: 'white', boxShadow: '0 6px 24px rgba(0,0,0,0.2)', borderRadius: 4 }}>
            <div style={{ padding: 12, maxHeight: '70vh', overflow: 'auto', fontFamily: 'Segoe UI, Arial' }}>
                <h3 style={{ marginTop: 0 }}>Select Calendars</h3>
                <div style={{ marginBottom: 8 }}>
                    <button onClick={() => { const all: Record<string, boolean> = {}; items.forEach(it => all[`${encodeURIComponent(it.webUrl)}::${it.id}`] = true); setSelected(all); }}>Select all</button>
                    <button style={{ marginLeft: 8 }} onClick={() => setSelected({})}>Clear</button>
                </div>
                {loading && <div>Discovering calendars...</div>}
                {!loading && error && <div style={{ color: 'darkred' }}>Discovery failed: {error}</div>}
                {!loading && !error && items.length === 0 && (
                    <div>
                        <div>No calendars found.</div>
                        <div style={{ marginTop: 8 }}>
                            <button onClick={() => {
                                // simple retry by re-running the effect: mutate context ref by forcing a tiny timeout change
                                setItems([]);
                                setLoading(true);
                                setTimeout(() => setLoading(false), 50);
                                // trigger discovery by calling discover via re-render: we call a no-op
                                // the effect depends on `context` which won't change; instead, call the discover logic directly here
                                (async () => {
                                    setLoading(true);
                                    setError(null);
                                    try {
                                        const siteRoot = context.pageContext.site && context.pageContext.site.absoluteUrl ? context.pageContext.site.absoluteUrl : context.pageContext.web.absoluteUrl;
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
                                        const discovered: IListEntry[] = [];
                                        for (const webUrl of webUrls) {
                                            try {
                                                const listsUrl = `${webUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title,BaseTemplate`;
                                                const lr = await context.spHttpClient.get(listsUrl, SPHttpClient.configurations.v1);
                                                if (!lr || !lr.ok) continue;
                                                const lj = await lr.json();
                                                const allLists = lj.value || [];
                                                for (const l of allLists) {
                                                    try {
                                                        const id = (l.Id || l.ID || l.id || '').toString();
                                                        const title = (l.Title || '').toString();
                                                        const base = l.BaseTemplate || 0;
                                                        if (base === 106 || title.toLowerCase().indexOf('event') !== -1 || title.toLowerCase().indexOf('calendar') !== -1) {
                                                            discovered.push({ id, title, webUrl });
                                                        }
                                                    } catch (inner) { }
                                                }
                                            } catch (e) { /* ignore */ }
                                        }
                                        setItems(discovered);
                                    } catch (e) {
                                        console.error('CalendarPicker (retry): discover failed', e);
                                        setError((e && e.message) || String(e));
                                    } finally {
                                        setLoading(false);
                                    }
                                })();
                            }}>Retry discovery</button>
                        </div>
                    </div>
                )}
                <ul style={{ listStyle: 'none', padding: 0 }}>
                    {items.map(it => {
                        const key = `${encodeURIComponent(it.webUrl)}::${it.id}`;
                        return (
                            <li key={key} style={{ marginBottom: 6 }}>
                                <label style={{ cursor: 'pointer' }}>
                                    <input type="checkbox" checked={!!selected[key]} onChange={() => toggle(key)} />
                                    <span style={{ marginLeft: 8 }}>{it.title} <small style={{ color: '#666' }}>— {it.webUrl}</small></span>
                                </label>
                            </li>
                        );
                    })}
                </ul>
                <div style={{ marginTop: 12 }}>
                    <button onClick={doSave} style={{ marginRight: 8 }}>Save</button>
                    <button onClick={onDismiss}>Cancel</button>
                </div>
                <div style={{ position: 'fixed', right: 12, top: 8 }}>
                    <button onClick={onDismiss} title="Close picker" style={{ background: 'transparent', border: 'none', fontSize: 18, cursor: 'pointer' }}>×</button>
                </div>
            </div>
        </div>
    );
};

export default CalendarPicker;
