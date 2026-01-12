import * as React from 'react';
import { ApplicationCustomizerContext } from '@microsoft/sp-application-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './TopNav.module.scss';

export interface ITopNavComponentProps {
  context: ApplicationCustomizerContext;
  hubRootWebUrl?: string;
  maxDepth?: number;
  cacheTtlMinutes?: number;
}

interface INavNode {
  Title: string;
  Url: string;
  Children?: { results?: INavNode[] } | INavNode[];
  // other fields omitted
}

export interface IMenuItem {
  title: string;
  url: string;
  children?: IMenuItem[];
}

const HARDCODED_HUB_URL = 'https://qchc2501.sharepoint.com/';

const TopNavComponent: React.FC<ITopNavComponentProps> = (props) => {
  const [menu, setMenu] = React.useState<IMenuItem[] | null>(null);
  const [loading, setLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | null>(null);

  React.useEffect(() => {
    const resolvedHubSiteUrl = (props.hubRootWebUrl && props.hubRootWebUrl.trim().length)
      ? props.hubRootWebUrl.replace(/\/$/, '')
      : (HARDCODED_HUB_URL && HARDCODED_HUB_URL.toString().trim().length)
        ? HARDCODED_HUB_URL.replace(/\/$/, '')
        : (props.context && props.context.pageContext && props.context.pageContext.web && props.context.pageContext.web.absoluteUrl)
          ? props.context.pageContext.web.absoluteUrl.replace(/\/$/, '')
          : '';

    if (!resolvedHubSiteUrl) {
      setError('Hub site URL not available');
      setLoading(false);
      return;
    }

    const maxDepth = props.maxDepth && Number.isInteger(props.maxDepth) ? props.maxDepth : 3;
    const expand = buildExpandString(maxDepth);

    // IMPORTANT: do NOT encode the expand string for SharePoint REST
    const endpoint = `${resolvedHubSiteUrl}/_api/web/Navigation/TopNavigationBar?$expand=${expand}`;
    const cacheKey = `spfx_topnav_${resolvedHubSiteUrl}`;
    const cacheTtlMinutes = props.cacheTtlMinutes && Number.isInteger(props.cacheTtlMinutes) ? props.cacheTtlMinutes : 10;

    const cached = readCache(cacheKey);
    if (cached) {
      setMenu(cached);
      setLoading(false);
      // background refresh
      fetchNav(endpoint, cacheKey, cacheTtlMinutes, props.context.spHttpClient, resolvedHubSiteUrl)
        .then(mapped => {
          // update cache and UI if changed
          setMenu(mapped);
        })
        .catch(e => console.warn('TopNav background refresh failed', e));
      return;
    }

    fetchNav(endpoint, cacheKey, cacheTtlMinutes, props.context.spHttpClient, resolvedHubSiteUrl)
      .then(mapped => {
        setMenu(mapped);
        setLoading(false);
      })
      .catch(err => {
        console.error('TopNav fetch error', err);
        setError('Failed to load navigation');
        setLoading(false);
      });

    // Re-run if hubRootWebUrl, maxDepth or cacheTtlMinutes change
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [props.hubRootWebUrl, props.maxDepth, props.cacheTtlMinutes]);

  if (loading) {
    return <div className={styles.topNavLoading || 'topNavLoading'}>Loading navigation…</div>;
  }
  if (error) {
    return <div className={styles.topNavError || 'topNavError'}>{error}</div>;
  }
  if (!menu || !menu.length) {
    return null;
  }

  // Determine hub host for link validation (used by isValidHref)
  const hubHost = (() => {
    try {
      const u = new URL((props.hubRootWebUrl && props.hubRootWebUrl.trim().length) ? props.hubRootWebUrl : HARDCODED_HUB_URL);
      return u.host.toLowerCase();
    } catch {
      return '';
    }
  })();

  return (
    <nav className={styles.topNav || 'topNav'}>
      <div className={styles.topNavInner || 'topNavInner'}>
        {/* brand removed per request */}
        <div className={styles.brand || 'brand'}>
          <img src="https://qchc2501.sharepoint.com/_api/siteiconmanager/getsitelogo?type=%271%27&hash=638883798660513472" alt="QCHC Logo" style={{ height: '40px' }} />
        </div>
        <div className={styles.menuWrap || 'menuWrap'}>
          <ul className={styles.topNavList || 'topNavList'}>
            {console.debug('TopNav rendering menu', menu)}
            {menu.map((item, idx) => <TopLevelMenuItem key={idx} item={item} hubHost={hubHost} />)}
          </ul>
        </div>
      </div>
    </nav>
  );
};

const TopLevelMenuItem: React.FC<{ item: IMenuItem; hubHost: string }> = ({ item, hubHost }) => {
  const hasChildren = !!(item.children && item.children.length);
  const validHref = isValidHref(item.url, hubHost);

  return (
    <li className={`${styles.topNavItem || 'topNavItem'} ${hasChildren ? 'has-children' : ''}`}>
      {validHref ? (
        <a className={styles.topNavLink || 'topNavLink'} href={item.url || '#'} data-nav-url={item.url || ''}>
          {item.title}
        </a>
      ) : (
        <span className={styles.topNavLink || 'topNavLink'} aria-disabled="true" data-nav-url={item.url || ''}>
          {item.title}
        </span>
      )}

      {hasChildren && (
        <div className={styles.dropdown || 'dropdown'} role="menu" aria-label={`${item.title} submenu`}>
          <ul className={styles.topNavList || 'topNavList'}>
            {item.children!.map((child, idx) => (
              <NestedMenuItem key={idx} item={child} hubHost={hubHost} />
            ))}
          </ul>
        </div>
      )}
    </li>
  );
};

const NestedMenuItem: React.FC<{ item: IMenuItem; hubHost: string }> = ({ item, hubHost }) => {
  const hasChildren = !!(item.children && item.children.length);
  const validHref = isValidHref(item.url, hubHost);

  return (
    <li className={`${styles.topNavItem || 'topNavItem'} ${hasChildren ? 'has-children' : ''}`}>
      {validHref ? (
        <a className={styles.topNavLink || 'topNavLink'} href={item.url || '#'} data-nav-url={item.url || ''}>
          {item.title}
        </a>
      ) : (
        <span className={styles.topNavLink || 'topNavLink'} aria-disabled="true" data-nav-url={item.url || ''}>
          {item.title}
        </span>
      )}

      {hasChildren && (
        <ul className={styles.subMenu || 'subMenu'}>
          {item.children!.map((child, idx) => (
            <NestedMenuItem key={idx} item={child} hubHost={hubHost} />
          ))}
        </ul>
      )}
    </li>
  );
};

/* ---------------- Helpers ---------------- */

function buildExpandString(maxDepth: number): string {
  const parts: string[] = [];
  for (let i = 1; i <= maxDepth; i++) {
    parts.push(Array(i).fill('Children').join('/'));
  }
  return parts.join(',');
}

function readCache(cacheKey: string): IMenuItem[] | null {
  try {
    const raw = sessionStorage.getItem(cacheKey);
    if (!raw) return null;
    const parsed = JSON.parse(raw);
    if (parsed && parsed.expires && parsed.data && parsed.expires > Date.now()) {
      return parsed.data as IMenuItem[];
    }
    sessionStorage.removeItem(cacheKey);
    return null;
  } catch (e) {
    console.warn('TopNav: sessionStorage read failed', e);
    return null;
  }
}

function writeCache(cacheKey: string, data: IMenuItem[], cacheTtlMinutes: number): void {
  try {
    const ttl = Math.max(1, cacheTtlMinutes) * 60 * 1000;
    const payload = { expires: Date.now() + ttl, data };
    sessionStorage.setItem(cacheKey, JSON.stringify(payload));
  } catch (e) {
    console.warn('TopNav: sessionStorage write failed', e);
  }
}

function normalizeUrl(rawUrl: string, hubSiteUrl: string): string {
  if (!rawUrl) return hubSiteUrl;
  if (/^https?:\/\//i.test(rawUrl)) return rawUrl;
  const hub = hubSiteUrl.replace(/\/$/, '');
  rawUrl = rawUrl.replace(/^~sitecollection/i, hub);
  rawUrl = rawUrl.replace(/^~site/i, hub);
  if (rawUrl.indexOf('/') === 0) {
    try {
      const u = new URL(hub);
      return `${u.protocol}//${u.host}${rawUrl}`;
    } catch (e) {
      return `${hub}${rawUrl}`;
    }
  }
  return `${hub}/${rawUrl.replace(/^\//, '')}`;
}

/**
 * Robust mapper: accepts Children as either { results: INavNode[] } or INavNode[].
 */
function mapNavNodeToMenuItem(node: INavNode, hubSiteUrl: string): IMenuItem {
  // children may be node.Children.results OR node.Children (array)
  let childrenNodes: INavNode[] = [];
  if (!node.Children) {
    childrenNodes = [];
  } else if (Array.isArray(node.Children)) {
    childrenNodes = node.Children as INavNode[];
  } else if ((node.Children as any).results && Array.isArray((node.Children as any).results)) {
    childrenNodes = (node.Children as any).results as INavNode[];
  } else {
    // fallback: try to treat as array-like
    try {
      childrenNodes = (Object.values(node.Children) as unknown) as INavNode[];
    } catch {
      childrenNodes = [];
    }
  }

  const children = childrenNodes.map(child => mapNavNodeToMenuItem(child, hubSiteUrl));
  const url = normalizeUrl((node as any).Url || '', hubSiteUrl);
  return {
    title: (node as any).Title || '',
    url,
    children: children.length ? children : undefined
  };
}

async function fetchNav(requestUrl: string, cacheKeyLocal: string, cacheTtlMinutes: number, spHttpClient: SPHttpClient, hubSiteUrl: string): Promise<IMenuItem[]> {
  const resp: SPHttpClientResponse = await spHttpClient.get(requestUrl, SPHttpClient.configurations.v1);
  if (!resp.ok) {
    throw new Error(`Failed to load navigation from ${requestUrl}: ${resp.status} ${resp.statusText}`);
  }
  const json = await resp.json();

  // Debug: inspect raw response to confirm nested Children shape
  console.debug('TopNavigationBar raw response', requestUrl, json);

  const nodes: INavNode[] = json.value || [];
  const mapped = nodes.map(n => mapNavNodeToMenuItem(n, hubSiteUrl));

  // Debug: inspect mapped structure
  console.debug('TopNav mapped menu', mapped);

  writeCache(cacheKeyLocal, mapped, cacheTtlMinutes);
  return mapped;
}

/**
 * Determine whether to render an <a href="..."> for the given URL.
 * Rules:
 * - URL must be an absolute http/https URL
 * - Exclude known placeholder hosts like 'linkless' or obviously invalid values
 * - Exclude javascript: and mailto: schemes
 */
function isValidHref(rawUrl: string | undefined, hubHost: string): boolean {
  if (!rawUrl) return false;
  const trimmed = rawUrl.trim();
  if (!trimmed) return false;

  const lower = trimmed.toLowerCase();
  if (lower === '#' || lower === 'javascript:void(0)' || lower === 'javascript:;' || lower.includes('linkless')) {
    return false;
  }
  if (lower.startsWith('mailto:') || lower.startsWith('tel:')) {
    return false;
  }

  try {
    const u = new URL(trimmed);
    if (u.protocol !== 'http:' && u.protocol !== 'https:') return false;
    // allow external https links; if you want to restrict to hubHost only, change this to compare hosts
    return true;
  } catch {
    // Not an absolute URL — treat as invalid here (normalizeUrl should have converted relative paths earlier)
    return false;
  }
}

export default TopNavComponent;