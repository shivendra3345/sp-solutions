import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';

import TopNav from './components/TopNav';
//import FooterInjector from './components/FooterInjector';
//import styles from './TopNavFooter.module.scss';

const LOG_SOURCE: string = 'TopNavFooterApplicationCustomizer';

export interface ITopNavFooterApplicationCustomizerProperties {
  topNavLinks?: string;   // JSON string or "Title|href,Title2|href2"
  footerHtml?: string;    // optional HTML to inject into footer
}

export default class TopNavFooterApplicationCustomizer
  extends BaseApplicationCustomizer<ITopNavFooterApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _footerContainerId = 'spfx-topnav-footer-injector';
  private _insertionContainerId = 'spfx-topnav-insert-after-search';
  private _observer: MutationObserver | null = null;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized TopNavFooterApplicationCustomizer`);

    // Inject critical hide style immediately to avoid header flash
    this._injectHideStyleEarly();

    // Render top placeholder (fallback)
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders.bind(this));
    this._renderPlaceholders();

    // Create footer container appended to document.body
    this._createFooterContainerAndRender();

    // Create and insert top nav after the page search box (override header visually)
    this._createAndInsertTopNavAfterSearch();

    return Promise.resolve();
  }
  /**
   * Inject a critical CSS rule as early as possible to hide the original header
   * before it renders. This prevents the header from appearing briefly.
   */
  private _hideStyleId = 'spfx-hide-header-critical-style';

  private _injectHideStyleEarly(): void {
    // If already injected, skip
    if (document.getElementById(this._hideStyleId)) {
      return;
    }

    // Narrow, explicit selectors for modern and classic headers
    const selectors = [
      '#spSiteHeader',
      '.spSiteHeader',
      '[data-automation-id="SiteHeader"]',
      '#SuiteNavPlaceHolder',
      '.od-TopBar',
      '.ms-compositeHeader',
      '.SPCommandBar'
    ];

    // Build a single rule string with !important and no transitions to avoid flicker
    const rule = `${selectors.join(', ')} { 
    display: none !important; 
    visibility: hidden !important; 
    height: 0 !important; 
    margin: 0 !important; 
    padding: 0 !important; 
    opacity: 0 !important; 
    transition: none !important; 
    pointer-events: none !important; 
  }`;

    const style = document.createElement('style');
    style.id = this._hideStyleId;
    style.appendChild(document.createTextNode(rule));

    // Insert as early as possible
    const head = document.head || document.getElementsByTagName('head')[0];
    if (head) {
      head.insertBefore(style, head.firstChild);
    } else {
      document.documentElement.appendChild(style);
    }

    // debug
    // eslint-disable-next-line no-console
    console.info('TopNavFooter: injected critical hide style to prevent header flash');
  }

  /**
   * Remove the critical hide style once our nav is inserted and visible.
   * Keeps the DOM clean and allows any other header logic to run if needed.
   */
  // private _removeHideStyleEarly(): void {
  //   const style = document.getElementById(this._hideStyleId);
  //   if (style) {
  //     style.remove();
  //     // eslint-disable-next-line no-console
  //     console.info('TopNavFooter: removed critical hide style');
  //   }
  // }
  private _renderPlaceholders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (this._topPlaceholder) {
        const topNavProps = {
          links: this._parseLinks(this.properties.topNavLinks)
        };

        ReactDOM.render(
          React.createElement(TopNav, topNavProps),
          this._topPlaceholder.domElement
        );
      }
    }
  }

  private _createFooterContainerAndRender(): void {
    let container = document.getElementById(this._footerContainerId);
    if (!container) {
      container = document.createElement('div');
      container.id = this._footerContainerId;
      document.body.appendChild(container);
    }

    //   const footerProps = {
    //     html: this.properties.footerHtml || ''
    //   };

    //   ReactDOM.render(React.createElement(FooterInjector, footerProps), container);
  }

  private _createAndInsertTopNavAfterSearch(): void {
    const temp = document.createElement('div');
    const topNavProps = { links: this._parseLinks(this.properties.topNavLinks) };
    ReactDOM.render(React.createElement(TopNav, topNavProps), temp);

    this._overrideHeaderAndInsertAfterSearch(temp);
  }

  /**
   * Safer header hide: only target narrow, explicit header selectors.
   * Avoid hiding ControlZone or other page content.
   */
  private _hideExistingHeader(): void {
    const selectors = [
      '#spSiteHeader',            // modern site header id
      '.spSiteHeader',            // modern site header class fallback
      '[data-automation-id="SiteHeader"]', // some tenants
      '#SuiteNavPlaceHolder',     // suite nav container
      '.od-TopBar',               // older top bar
      '.ms-compositeHeader',      // composite header
      '.SPCommandBar',            // command bar (if present)
      '.ms-HubNav',       // hub nav (if present)
      '.ms-HorizontalNav'

    ];

    const hidden: string[] = [];
    selectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        (el as HTMLElement).classList.add('spfx-hidden-original-header');
        hidden.push(sel);
      });
    });

    if (!document.getElementById('spfx-hide-header-style')) {
      const style = document.createElement('style');
      style.id = 'spfx-hide-header-style';
      style.innerHTML = `
        .spfx-hidden-original-header { display: none !important; visibility: hidden !important; height: 0 !important; margin: 0 !important; padding: 0 !important; }
        #${this._insertionContainerId} { display: block; }
      `;
      document.head.appendChild(style);
    }

    // eslint-disable-next-line no-console
    console.info('TopNavFooter: hid header selectors:', hidden);
  }

  private _insertTopNavAfterSearch(topNavNode: HTMLElement): void {
    const searchSelectors = [
      'div[data-automation-id="SearchBox"]',
      'div[data-automation-id="SearchBoxContainer"]',
      '#SearchBoxContainer',
      //'.ms-SearchBox',
      'div[role="search"]',
      'input[title="Search"]'
    ];

    const findSearchEl = (): HTMLElement | null => {
      for (const sel of searchSelectors) {
        const el = document.querySelector(sel) as HTMLElement | null;
        if (el) {
          // eslint-disable-next-line no-console
          console.info('TopNavFooter: found search element using selector', sel, el);
          return el;
        }
      }
      const fallback = document.querySelector('#suiteNavBox') as HTMLElement | null;
      // eslint-disable-next-line no-console
      console.info('TopNavFooter: fallback search element', !!fallback);
      return fallback;
    };

    const insertAfter = (reference: HTMLElement, node: HTMLElement) => {
      if (reference.parentNode) {
        reference.parentNode.insertBefore(node, reference.nextSibling);
      } else {
        document.body.appendChild(node);
      }
    };

    const attemptInsert = () => {
      const searchEl = findSearchEl();
      if (searchEl) {
        if (!document.getElementById(this._insertionContainerId)) {
          const container = document.createElement('div');
          container.id = this._insertionContainerId;
          container.className = 'spfx-topnav-insert';
          while (topNavNode.firstChild) {
            container.appendChild(topNavNode.firstChild);
          }
          insertAfter(searchEl, container);
          // eslint-disable-next-line no-console
          console.info('TopNavFooter: inserted top nav after search element');
        }
        if (this._observer) {
          this._observer.disconnect();
          this._observer = null;
        }
        return true;
      }
      return false;
    };

    if (attemptInsert()) return;

    const root = document.body;
    this._observer = new MutationObserver(() => {
      if (attemptInsert()) {
        if (this._observer) {
          this._observer.disconnect();
          this._observer = null;
        }
      }
    });

    this._observer.observe(root, { childList: true, subtree: true });

    setTimeout(() => {
      if (this._observer) {
        this._observer.disconnect();
        this._observer = null;
        // eslint-disable-next-line no-console
        console.warn('TopNavFooter: stopped observing DOM for search element (timeout)');
      }
    }, 10000);
  }

  private _overrideHeaderAndInsertAfterSearch(topNavElement: HTMLElement): void {
    this._hideExistingHeader();
    this._insertTopNavAfterSearch(topNavElement);
  }

  private _parseLinks(raw?: string): { title: string; href: string }[] {
    if (!raw) return [
      { title: 'Home', href: '/' },
      { title: 'News', href: '/SitePages/News.aspx' },
      { title: 'Directory', href: '/SitePages/Directory.aspx' }
    ];

    try {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed)) return parsed;
    } catch (e) {
      // ignore
    }

    return raw.split(',').map(s => {
      const parts = s.split('|').map(p => p.trim());
      return { title: parts[0] || parts.join('|'), href: parts[1] || '#' };
    });
  }

  private _onDispose(): void {
    // placeholder disposed
  }

  @override
  public onDispose(): void {
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement);
    }

    const footerContainer = document.getElementById(this._footerContainerId);
    if (footerContainer) {
      ReactDOM.unmountComponentAtNode(footerContainer);
      footerContainer.remove();
    }

    const inserted = document.getElementById(this._insertionContainerId);
    if (inserted) {
      ReactDOM.unmountComponentAtNode(inserted);
      inserted.remove();
    }

    const hideStyle = document.getElementById('spfx-hide-header-style');
    if (hideStyle) {
      hideStyle.remove();
    }

    if (this._observer) {
      this._observer.disconnect();
      this._observer = null;
    }
  }
}