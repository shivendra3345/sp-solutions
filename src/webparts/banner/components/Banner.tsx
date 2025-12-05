import * as React from 'react';
import styles from './Banner.module.scss';
import type { IBannerProps } from './IBannerProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import HeroBanner from './Banner/HeroBanner';


export default class Banner extends React.Component<IBannerProps, {}> {
  public render(): React.ReactElement<IBannerProps> {
    const {
      hasTeamsContext,
      context
    } = this.props;

    return (
      <section id='customSec' className={`${styles.banner} 'customSec' ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <HeroBanner context={context} heroImageUrl='https://qchc2501.sharepoint.com/SiteAssets/Images/banner.jpg' />
        </div>
      </section>
    );
  }
}
