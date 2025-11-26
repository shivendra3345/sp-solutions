import * as React from 'react';
import styles from '../TopNavFooter.module.scss';

export interface IFooterInjectorProps {
  html?: string;
}

const FooterInjector: React.FC<IFooterInjectorProps> = ({ html = '' }) => {
  return (
    <div className={styles.injectedFooterWrapper} role="contentinfo" aria-label="Site footer">
      <div
        className={styles.injectedFooterInner}
        dangerouslySetInnerHTML={{ __html: html || defaultFooterHtml() }}
      />
    </div>
  );
};

function defaultFooterHtml(): string {
  return `
    <div class="footer-content">
      <div class="footer-left">© ${new Date().getFullYear()} Your Company</div>
      <div class="footer-right">
        <a href="/privacy">Privacy</a> · <a href="/terms">Terms</a>
      </div>
    </div>
  `;
}

export default FooterInjector;