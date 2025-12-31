import * as React from 'react';
import * as ReactDOM from 'react-dom';
import styles from './Banner.module.scss';
import type { IBannerProps } from './IBannerProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import HeroBanner from './Banner/HeroBanner';


export default class Banner extends React.Component<IBannerProps, {}> {
  private HEADER_ID = 'sp-solutions-header-extension';

  private ensureHeaderHost(enabled?: boolean) {
    if (typeof document === 'undefined') return;
    const id = this.HEADER_ID;
    let host = document.getElementById(id);
    if (enabled) {
      if (!host) {
        host = document.createElement('div');
        host.id = id;
        host.setAttribute('role', 'banner');
        host.style.position = 'fixed';
        host.style.top = '0';
        host.style.left = '0';
        host.style.right = '0';
        host.style.zIndex = '99998';
        host.style.background = 'transparent';
        document.body.appendChild(host);
      }
      // mark as enabled and ensure visible
      host.setAttribute('data-enabled', 'true');
      host.style.display = '';
      try {
        console.debug('Banner: dispatching sp-solutions-header-toggle { enabled: true }');
        // persist preference so extensions can read it on page load
        try { localStorage.setItem('sp-solutions-header-enabled', 'true'); } catch (e) { /* ignore */ }
        window.dispatchEvent(new CustomEvent('sp-solutions-header-toggle', { detail: { enabled: true } }));
      } catch (e) { /* ignore */ }
    } else {
      if (host) {
        // mark as disabled and hide; don't remove element to allow other scripts to observe attribute
        host.setAttribute('data-enabled', 'false');
        host.style.display = 'none';
        host.style.pointerEvents = 'none';
        host.setAttribute('aria-hidden', 'true');
        // attempt to unmount any React content placed by other scripts into this host
        try { ReactDOM.unmountComponentAtNode(host); } catch (ux) { /* ignore */ }
        try {
          console.debug('Banner: dispatching sp-solutions-header-toggle { enabled: false }');
          // persist preference so extensions can read it on page load
          try { localStorage.setItem('sp-solutions-header-enabled', 'false'); } catch (e) { /* ignore */ }
          window.dispatchEvent(new CustomEvent('sp-solutions-header-toggle', { detail: { enabled: false } }));
        } catch (e) { /* ignore */ }
      }
    }
  }
  public render(): React.ReactElement<IBannerProps> {
    // ensure header host exists/removed based on prop each render (lightweight)
    try {
      const enabled = this.props.enableHeaderExtension !== false;
      // debug: log current requested header extension state
      try { console.debug('Banner: enableHeaderExtension prop value ->', this.props.enableHeaderExtension, 'resolvedEnabled ->', enabled); } catch (e) { /* ignore */ }
      this.ensureHeaderHost(enabled);
    } catch (e) { /* swallow */ }
    const {
      hasTeamsContext,

    } = this.props;
    // heroImageUrl = 'https://qchc2501.sharepoint.com/SiteAssets/Images/banner.jpg' 
    return (
      <section id='customSec' className={`${styles.banner} 'customSec' ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <HeroBanner {...this.props} />
        </div>
      </section>
    );
  }
}
