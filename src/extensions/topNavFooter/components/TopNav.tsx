import * as React from 'react';
import { useEffect, useState, useRef } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './TopNav.module.scss';

export interface INavNode {
  Id: string;
  Title: string;
  Url: string;
  IsExternal?: boolean;
  Children?: INavNode[];
}

export interface ITopNavProps {
  context: any; // ExtensionContext / WebPartContext
  hubRootWebUrl?: string; // optional: will be hardcoded below per request
  maxDepth?: number; // how many levels to render (default 2)
  cacheTtlSeconds?: number; // cache TTL in seconds (default 300)
}

const DEFAULT_CACHE_TTL = 300;

const TopNav: React.FC<ITopNavProps> = ({
  context,
  hubRootWebUrl,
  maxDepth = 2,
  cacheTtlSeconds = DEFAULT_CACHE_TTL
}) => {
  const [nav, setNav] = useState<INavNode[]>([]);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);
  const [mobileOpen, setMobileOpen] = useState<boolean>(false);
  const mountedRef = useRef(true);

  // Keep a ref to the currently opened dropdown (for keyboard/escape handling)
  const openDropdownRef = useRef<HTMLLIElement | null>(null);

  useEffect(() => {
    mountedRef.current = true;
    setLoading(true);
    setError(null);

    // Hardcoded hub URL as requested
    // eslint-disable-next-line no-param-reassign
    hubRootWebUrl = 'https://lcor1.sharepoint.com/sites/communication/';

    const baseUrl = (hubRootWebUrl && hubRootWebUrl.toString().trim().length)
      ? hubRootWebUrl.replace(/\/$/, '')
      : (context && context.pageContext && context.pageContext.web && context.pageContext.web.absoluteUrl)
        ? context.pageContext.web.absoluteUrl.replace(/\/$/, '')
        : '';

    if (!baseUrl) {
      setError('Unable to determine site URL for navigation.');
      setLoading(false);
      // eslint-disable-next-line no-console
      console.error('TopNav: baseUrl not found (hubRootWebUrl or context.pageContext.web.absoluteUrl)');
      return;
    }

    const endpoint = `${baseUrl}/_api/web/Navigation/TopNavigationBar?$expand=Children`;
    const cacheKey = `spfx_topnav_${baseUrl}`;

    // Try session cache first
    try {
      const raw = sessionStorage.getItem(cacheKey);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && parsed.expires && parsed.data && parsed.expires > Date.now()) {
          setNav(parsed.data);
          setLoading(false);
          // background refresh
          fetchNav(endpoint, cacheKey, false);
          return;
        } else {
          sessionStorage.removeItem(cacheKey);
        }
      }
    } catch (e) {
      // ignore storage errors
      // eslint-disable-next-line no-console
      console.warn('TopNav: sessionStorage read failed', e);
    }

    // No valid cache — fetch and set
    fetchNav(endpoint, cacheKey, true);

    return () => {
      mountedRef.current = false;
      // close any open dropdown on unmount
      openDropdownRef.current = null;
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [context, hubRootWebUrl, maxDepth, cacheTtlSeconds]);

  function fetchNav(endpoint: string, cacheKey: string, setUi: boolean) {
    if (!context || !context.spHttpClient) {
      if (mountedRef.current) {
        setError('SPHttpClient not available on context.');
        setLoading(false);
      }
      // eslint-disable-next-line no-console
      console.error('TopNav: context.spHttpClient missing');
      return;
    }

    // eslint-disable-next-line no-console
    console.info('TopNav: fetching navigation from', endpoint);

    context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => {
        if (!res.ok) {
          throw new Error(`Failed to load navigation: ${res.status} ${res.statusText}`);
        }
        return res.json();
      })
      .then((data: any) => {
        if (!mountedRef.current) return;
        const nodes: INavNode[] = (data && data.value ? data.value : []).map((n: any) => mapNode(n));
        if (setUi) {
          setNav(nodes);
          setLoading(false);
        }
        try {
          const payload = { expires: Date.now() + (cacheTtlSeconds * 1000), data: nodes };
          sessionStorage.setItem(cacheKey, JSON.stringify(payload));
        } catch (e) {
          // ignore storage errors
        }
      })
      .catch((err: any) => {
        if (!mountedRef.current) return;
        setError(err && err.message ? err.message : String(err));
        setLoading(false);
        // eslint-disable-next-line no-console
        console.error('TopNav fetch error', err);
      });
  }

  function mapNode(n: any): INavNode {
    const id = n && (n.Id || n.Id === 0) ? String(n.Id) : Math.random().toString(36).substr(2, 9);
    const title = (n && (n.Title || n.Title === '')) ? String(n.Title) : '';
    let url = '#';

    if (n) {
      if (typeof n.Url === 'string' && n.Url.trim().length) url = n.Url;
      else if (n.Url && typeof n.Url === 'object') url = (n.Url.Url && typeof n.Url.Url === 'string') ? n.Url.Url : (n.Url && n.Url.toString ? String(n.Url) : '#');
      else if (n.Url && n.Url.Url) url = n.Url.Url;
    }

    const isExternal = !!(url && /^https?:\/\//i.test(url) && url.indexOf(window.location.hostname) === -1);

    const node: INavNode = { Id: id, Title: title, Url: url || '#', IsExternal: isExternal, Children: [] };

    const childrenRaw = n && n.Children ? n.Children : null;
    const childrenArray = Array.isArray(childrenRaw) ? childrenRaw : (childrenRaw && childrenRaw.results ? childrenRaw.results : []);
    if (childrenArray && childrenArray.length) node.Children = childrenArray.map((c: any) => mapNode(c));

    return node;
  }

  // keyboard handlers for accessibility and navigation
  const onKeyDownItem = (e: React.KeyboardEvent, li: HTMLLIElement | null) => {
    if (!li) return;
    const key = e.key;
    const hasChildren = !!li.querySelector(`.${styles.dropdown}`);
    if (key === 'Enter' || key === ' ') {
      if (hasChildren) {
        e.preventDefault();
        const expanded = li.getAttribute('data-expanded') === 'true';
        li.setAttribute('data-expanded', (!expanded).toString());
        if (!expanded) {
          li.classList.add(styles.topNavItemHasChildren);
          openDropdownRef.current = li;
        } else {
          li.classList.remove(styles.topNavItemHasChildren);
          openDropdownRef.current = null;
        }
      }
    } else if (key === 'Escape') {
      if (hasChildren) {
        li.setAttribute('data-expanded', 'false');
        li.classList.remove(styles.topNavItemHasChildren);
        openDropdownRef.current = null;
      }
    } else if (key === 'ArrowDown') {
      if (hasChildren) {
        e.preventDefault();
        const firstChildLink = li.querySelector(`.${styles.dropdown} a`) as HTMLElement | null;
        if (firstChildLink) firstChildLink.focus();
      }
    } else if (key === 'ArrowRight') {
      const next = li.nextElementSibling as HTMLElement | null;
      if (next) {
        const link = next.querySelector(`.${styles.topNavLink}`) as HTMLElement | null;
        if (link) link.focus();
      }
    } else if (key === 'ArrowLeft') {
      const prev = li.previousElementSibling as HTMLElement | null;
      if (prev) {
        const link = prev.querySelector(`.${styles.topNavLink}`) as HTMLElement | null;
        if (link) link.focus();
      }
    }
  };

  // close any open dropdown when clicking outside
  useEffect(() => {
    const onDocClick = (ev: MouseEvent) => {
      const target = ev.target as HTMLElement | null;
      if (!target) return;
      const open = openDropdownRef.current;
      if (open && !open.contains(target)) {
        open.setAttribute('data-expanded', 'false');
        open.classList.remove(styles.topNavItemHasChildren);
        openDropdownRef.current = null;
      }
    };

    document.addEventListener('click', onDocClick, true);
    return () => document.removeEventListener('click', onDocClick, true);
  }, []);

  const toggleMobile = () => setMobileOpen(v => !v);

  const renderMenu = (items: INavNode[], depth = 1) => {
    if (!items || !items.length) return null;
    return (
      <ul
        className={styles.topNavList}
        role={depth === 1 ? 'menubar' : 'menu'}
        aria-label={depth === 1 ? 'Top navigation' : 'Sub menu'}
      >
        {items.map(item => {
          let liRef: HTMLLIElement | null = null;
          const setLiRef = (el: HTMLLIElement | null) => { liRef = el; };

          return (
            <li
              key={item.Id}
              className={styles.topNavItem}
              role="none"
              ref={setLiRef}
              tabIndex={0}
              onKeyDown={(e) => onKeyDownItem(e, liRef)}
              data-expanded="false"
            >
              <a
                href={item.Url || '#'}
                className={styles.topNavLink}
                role="menuitem"
                target={item.IsExternal ? '_blank' : undefined}
                rel={item.IsExternal ? 'noopener noreferrer' : undefined}
              >
                {item.Title}
              </a>

              {item.Children && item.Children.length > 0 && depth < maxDepth && (
                <div className={styles.dropdown} role="presentation" aria-hidden="true">
                  {renderMenu(item.Children, depth + 1)}
                </div>
              )}
            </li>
          );
        })}
      </ul>
    );
  };

  return (
    <nav className={styles.topNav} aria-label="Site top navigation">
      <div className={styles.topNavInner}>
        {/* <div className={styles.brand} aria-hidden="true">THE HUB</div> */}

        <div className={styles.menuWrap}>
          <button
            className={styles.mobileToggle}
            aria-expanded={mobileOpen}
            aria-controls="spfx-mobile-topnav"
            onClick={toggleMobile}
            type="button"
          >
            <span className={styles.srOnly}>Toggle navigation</span>
            ☰
          </button>

          {loading && <div className={styles.loading}>Loading navigation…</div>}
          {error && <div className={styles.error}>Navigation error: {error}</div>}

          {!loading && !error && (
            <>
              <div className={styles.desktopMenu} aria-hidden={mobileOpen ? 'true' : 'false'}>
                {renderMenu(nav)}
              </div>

              <div
                id="spfx-mobile-topnav"
                className={`${styles.mobileMenu} ${mobileOpen ? styles.mobileMenuOpen : ''}`}
                aria-hidden={!mobileOpen}
              >
                {renderMenu(nav)}
              </div>
            </>
          )}
        </div>
      </div>
    </nav>
  );
};

export default TopNav;