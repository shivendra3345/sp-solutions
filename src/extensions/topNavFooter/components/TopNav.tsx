import * as React from 'react';
import { Link } from '@fluentui/react';
import styles from '../TopNavFooter.module.scss';

export interface ITopNavProps {
  links?: { title: string; href: string }[];
}

const TopNav: React.FC<ITopNavProps> = ({ links = [] }) => {
  return (
    <div className={styles.topNav} role="navigation" aria-label="Top navigation">
      <div className={styles.topNavInner}>
        <div className={styles.brand}>THE HUB</div>
        <nav className={styles.nav} aria-label="Primary">
          {links.map((l, i) => (
            <Link key={i} href={l.href} className={styles.navLink}>
              {l.title}
            </Link>
          ))}
        </nav>
      </div>
    </div>
  );
};

export default TopNav;