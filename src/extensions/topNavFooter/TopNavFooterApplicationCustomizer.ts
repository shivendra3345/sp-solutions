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

export interface ITopNavFooterApplicationCustomizerProperties {
  hubRootWebUrl?: string;
  footerHtml?: string;
}

const LOG_SOURCE: string = 'TopNavFooterApplicationCustomizer';

export default class TopNavFooterApplicationCustomizer
  extends BaseApplicationCustomizer<ITopNavFooterApplicationCustomizerProperties> {

  private _topPlaceholder: PlaceholderContent | undefined;
  private _insertionContainerId = 'spfx-topnav-insert-after-search';
  private _footerContainerId = 'spfx-injected-footer-container';
  private _criticalHideStyleId = 'spfx-hide-header-critical-style';
  private _reapplyObserverKey = '__spfx_header_reapply_observer__';
  private _reapplyObserver: MutationObserver | null = null;
  private _observer: MutationObserver | null = null;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized TopNavFooterApplicationCustomizer`);
    // If the user previously disabled the header via the Banner web part, persist
    // that preference in localStorage. Read it here and skip initial injection
    // if the stored preference is 'false'.
    let preferEnabled: boolean = true;
    try {
      const pref = (typeof window !== 'undefined' && localStorage.getItem('sp-solutions-header-enabled')) || undefined;
      if (pref === 'false') preferEnabled = false;
    } catch (e) { /* ignore */ }

    if (!preferEnabled) {
      try { console.debug('TopNavFooter: user preference indicates header extension disabled on load; skipping injection'); } catch (e) { }
      // ensure any previous injected UI is removed
      try { this._disableCustomHeader(); } catch (e) { /* ignore */ }
    } else {
      // Inject critical hide style as early as possible to avoid header flash
      this._injectHideStyleEarly();

      // Render into Top placeholder as fallback
      this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceholders.bind(this));
      this._renderPlaceholders();

      // Create footer container appended to body and render footer
      this._createFooterContainerAndRender();

      // Create TopNav element and attempt to insert after search box
      this._createAndInsertTopNavAfterSearch();
    }

    // Listen for header toggle events from the Banner web part so we can show/hide
    // the custom header and restore the hub header when disabled.
    try {
      window.addEventListener('sp-solutions-header-toggle', (ev: any) => {
        const enabled = !!(ev && ev.detail && ev.detail.enabled);
        try { console.debug('TopNavFooter: received sp-solutions-header-toggle ->', enabled); } catch (e) { }
        if (enabled) {
          // re-enable custom header
          this._enableCustomHeader();
        } else {
          // disable custom header and restore original header elements
          this._disableCustomHeader();
        }
      });
    } catch (e) {
      // ignore if events fail
    }

    return Promise.resolve();
  }

  private _renderPlaceholders(): void {
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      if (this._topPlaceholder && this._topPlaceholder.domElement) {
        const topNavProps = {
          context: this.context,
          hubRootWebUrl: this.properties.hubRootWebUrl
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

    // const footerProps = {
    //   html: this.properties.footerHtml || ''
    // };

    //  ReactDOM.render(React.createElement(FooterInjector, footerProps), container);
  }

  private _createAndInsertTopNavAfterSearch(): void {
    const temp = document.createElement('div');
    const topNavProps = {
      context: this.context,
      hubRootWebUrl: this.properties.hubRootWebUrl
    };

    ReactDOM.render(React.createElement(TopNav, topNavProps), temp);

    this._insertTopNavAfterSearch(temp);
  }

  private _insertTopNavAfterSearch(topNavNode: HTMLElement): void {
    const searchSelectors = [
      'div[data-automation-id="SearchBox"]',
      'div[data-automation-id="SearchBoxContainer"]',
      '#SearchBoxContainer',
      // '.ms-SearchBox',
      'div[role="search"]',
      'input[title="Search"]',
      '#suiteNavBox'
    ];

    const findSearchEl = (): HTMLElement | null => {
      for (const sel of searchSelectors) {
        const el = document.querySelector(sel) as HTMLElement | null;
        if (el) {
          return el;
        }
      }
      return null;
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

          this._applyHideClassToHeaderElements();
          //  this._removeHideStyleEarly();
          this._startHeaderReapplyObserver();

          return true;
        }
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
        if (!document.getElementById(this._insertionContainerId)) {
          // this._removeHideStyleEarly();
          // eslint-disable-next-line no-console
          console.warn('TopNavFooter: stopped observing DOM for search element (timeout) — removed hide style');
        }
      }
    }, 10000);
  }

  private _injectHideStyleEarly(): void {
    if (document.getElementById(this._criticalHideStyleId)) return;

    const selectors = [
      '#spSiteHeader',
      '.spSiteHeader',
      '[data-automation-id="SiteHeader"]',
      '#SuiteNavPlaceHolder',
      '.od-TopBar',
      '.ms-compositeHeader',
      '.SPCommandBar',
      '#DeltaPlaceHolderPageTitleInTitleArea',
      '.ms-HubNav',
      '.ms-HubNav-enhancedMegaMenu',
      '.ms-HorizontalNav'
    ];

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
    style.id = this._criticalHideStyleId;
    style.appendChild(document.createTextNode(rule));

    const head = document.head || document.getElementsByTagName('head')[0];
    if (head) {
      head.insertBefore(style, head.firstChild);
    } else {
      document.documentElement.appendChild(style);
    }

    // eslint-disable-next-line no-console
    console.info('TopNavFooter: injected critical hide style');
  }

  private _removeHideStyleEarly(): void {
    const style = document.getElementById(this._criticalHideStyleId);
    if (style) {
      style.remove();
      // eslint-disable-next-line no-console
      console.info('TopNavFooter: removed critical hide style');
    }
  }

  private _disableCustomHeader(): void {
    try {
      console.debug('TopNavFooter: performing forced disable cleanup');
    } catch (e) { }

    // Unmount and remove inserted container if present
    const inserted = document.getElementById(this._insertionContainerId);
    if (inserted) {
      try { ReactDOM.unmountComponentAtNode(inserted); } catch (e) { /* ignore */ }
      inserted.remove();
      try { console.debug('TopNavFooter: removed insertion container', this._insertionContainerId); } catch (e) { }
    }

    // Remove footer container
    const footerContainer = document.getElementById(this._footerContainerId);
    if (footerContainer) {
      try { ReactDOM.unmountComponentAtNode(footerContainer); } catch (e) { /* ignore */ }
      footerContainer.remove();
      try { console.debug('TopNavFooter: removed footer container', this._footerContainerId); } catch (e) { }
    }

    // remove hide styles and un-hide original header elements
    const hideStyle = document.getElementById('spfx-hide-header-style');
    if (hideStyle) {
      hideStyle.remove();
      try { console.debug('TopNavFooter: removed style spfx-hide-header-style'); } catch (e) { }
    }
    // remove the early critical hide style as well
    const critical = document.getElementById(this._criticalHideStyleId);
    if (critical) {
      critical.remove();
      try { console.debug('TopNavFooter: removed critical hide style', this._criticalHideStyleId); } catch (e) { }
    }
    this._removeHideStyleEarly();

    // remove any inserted nodes with known class names that may still be present
    document.querySelectorAll('.spfx-topnav-insert, .spfx-topnav-root, .spfx-topnav-inserted').forEach(el => {
      try { el.remove(); } catch (e) { /* ignore */ }
    });

    // unhide original header elements
    document.querySelectorAll('.spfx-hidden-original-header').forEach(el => {
      try { (el as HTMLElement).classList.remove('spfx-hidden-original-header'); } catch (e) { }
    });

    // try to clear inline styles that may hide headers (be conservative)
    const selectors = [
      '#spSiteHeader', '.spSiteHeader', '[data-automation-id="SiteHeader"]', '#SuiteNavPlaceHolder', '.od-TopBar', '.ms-compositeHeader', '.SPCommandBar', '#DeltaPlaceHolderPageTitleInTitleArea', '.ms-HubNav', '.ms-HubNav-enhancedMegaMenu', '.ms-HorizontalNav'
    ];
    selectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        try {
          const e = el as HTMLElement;
          e.style.removeProperty('display');
          e.style.removeProperty('visibility');
          e.style.removeProperty('height');
          e.style.removeProperty('margin');
          e.style.removeProperty('padding');
          e.style.removeProperty('opacity');
          e.style.removeProperty('pointer-events');
        } catch (err) { /* ignore */ }
      });
    });

    // stop observers that re-apply hide styles
    this._stopHeaderReapplyObserver();

    // also remove the top placeholder content if we used it
    if (this._topPlaceholder && this._topPlaceholder.domElement) {
      try { ReactDOM.unmountComponentAtNode(this._topPlaceholder.domElement); } catch (e) { /* ignore */ }
      try { console.debug('TopNavFooter: unmounted top placeholder'); } catch (e) { }
    }

    try { console.debug('TopNavFooter: forced cleanup complete'); } catch (e) { }
    return;


  }

  private _enableCustomHeader(): void {
    // Re-inject critical hide style and attempt to re-insert top nav
    try {
      this._injectHideStyleEarly();
      this._renderPlaceholders();
      this._createAndInsertTopNavAfterSearch();
    } catch (e) {
      console.warn('TopNavFooter: enableCustomHeader failed', e);
    }
  }

  // private _removeHideStyleEarly(): void {
  //   const style = document.getElementById(this._criticalHideStyleId);
  //   if (style) {
  //     style.remove();
  //     // eslint-disable-next-line no-console
  //     console.info('TopNavFooter: removed critical hide style');
  //   }
  // }

  private _applyHideClassToHeaderElements(): void {
    const selectors = [
      '#spSiteHeader',
      '.spSiteHeader',
      '[data-automation-id="SiteHeader"]',
      '#SuiteNavPlaceHolder',
      '.od-TopBar',
      '.ms-compositeHeader',
      '.SPCommandBar',
      '#DeltaPlaceHolderPageTitleInTitleArea',
      '.ms-HubNav',
      '.ms-HubNav-enhancedMegaMenu',
      '.ms-HorizontalNav'
    ];

    selectors.forEach(sel => {
      document.querySelectorAll(sel).forEach(el => {
        const e = el as HTMLElement;
        if (e && e.offsetHeight > 0) {
          e.classList.add('spfx-hidden-original-header');
        }
      });
    });

    if (!document.getElementById('spfx-hide-header-style')) {
      const style = document.createElement('style');
      style.id = 'spfx-hide-header-style';
      style.innerHTML = `
        .spfx-hidden-original-header { display: none !important; visibility: hidden !important; height: 0 !important; margin: 0 !important; padding: 0 !important; opacity: 0 !important; pointer-events: none !important; }
      `;
      document.head.appendChild(style);
    }
  }

  private _startHeaderReapplyObserver(): void {
    if ((window as any)[this._reapplyObserverKey]) return;

    const hideFn = () => {
      const selectors = [
        '#spSiteHeader',
        '.spSiteHeader',
        '[data-automation-id="SiteHeader"]',
        '#SuiteNavPlaceHolder',
        '.od-TopBar',
        '.ms-compositeHeader',
        '.SPCommandBar',
        '#DeltaPlaceHolderPageTitleInTitleArea'
      ];
      selectors.forEach(sel => {
        document.querySelectorAll(sel).forEach(el => {
          const e = el as HTMLElement;
          if (e && e.offsetHeight > 0) {
            e.classList.add('spfx-hidden-original-header');
          }
        });
      });
    };

    hideFn();

    const observer = new MutationObserver(mutations => {
      for (const m of mutations) {
        if (m.type === 'childList' || m.type === 'attributes') {
          hideFn();
          break;
        }
      }
    });

    observer.observe(document.body, { childList: true, subtree: true, attributes: true, attributeFilter: ['class', 'style', 'id'] });

    (window as any)[this._reapplyObserverKey] = observer;
    this._reapplyObserver = observer;

    // eslint-disable-next-line no-console
    console.info('TopNavFooter: started header reapply observer');
  }

  private _stopHeaderReapplyObserver(): void {
    const obs = (window as any)[this._reapplyObserverKey];
    if (obs && typeof obs.disconnect === 'function') {
      obs.disconnect();
      delete (window as any)[this._reapplyObserverKey];
    }
    if (this._reapplyObserver) {
      this._reapplyObserver.disconnect();
      this._reapplyObserver = null;
    }
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

    //  this._removeHideStyleEarly();
    this._stopHeaderReapplyObserver();

    if (this._observer) {
      this._observer.disconnect();
      this._observer = null;
    }
  }
}