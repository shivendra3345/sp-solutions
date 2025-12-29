import * as React from 'react';
import { useEffect, useState } from 'react';
import {
  Persona,
  PersonaSize,
  //IconButton,
  // Stack,
  Pivot,
  PivotItem
} from '@fluentui/react';
import styles from './HeroBanner.module.scss';
import type { IBannerProps } from '../IBannerProps';
import WelcomeNewJoiners from './WelcomeNewJoiners';
import EventsCarousel from './EventsCarousel';
import InfoTile from './InfoTile';
import DataService, { } from '../../services/DataService';
import NewsList from '../news/NewsList';
import AnnouncementList from '../announcement/AnnouncementList';
import CombinedFeed from './CombinedFeed';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IBannerItem {
  Id: number;
  Title: string;
  Active: boolean;
  AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
}

type TabKey = 'all' | 'news' | 'announcements';

const HeroBanner: React.FC<IBannerProps> = ({ context, showWelcome, excludedSites }) => {
  const [userName, setUserName] = useState<string>('User');
  //const [featured, setFeatured] = useState<IFeaturedItem[]>([]);
  const [activeTab, setActiveTab] = useState<TabKey>('all');
  const [bannerImages, setBannerImages] = useState<string[]>([]);
  const [currentIndex, setCurrentIndex] = useState<number>(0);

  useEffect(() => {
    const svc = new DataService(context);
    svc.getCurrentUserDisplayName().then(name => setUserName(name));
    // svc.getFeaturedItems().then(items => setFeatured(items));

    // Fetch banner images from SharePoint list "Banner"
    const url =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Banner')/items` +
      `?$select=Id,Title,Active,AttachmentFiles&$expand=AttachmentFiles&$top=3`;

    context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data) => {
        const items: IBannerItem[] = data.value || [];
        const activeImages: string[] = [];
        items.forEach(item => {
          if (item.Active && item.AttachmentFiles && item.AttachmentFiles.length > 0) {
            activeImages.push(item.AttachmentFiles[0].ServerRelativeUrl);
          }
        });
        setBannerImages(activeImages);
      })
      .catch(err => console.error("Error fetching banner images:", err));
  }, [context]);

  // Carousel auto-advance
  useEffect(() => {
    if (bannerImages.length > 1) {
      const interval = setInterval(() => {
        setCurrentIndex(prev => (prev + 1) % bannerImages.length);
      }, 5000); // 5s
      return () => clearInterval(interval);
    }
  }, [bannerImages]);

  /*  function getTags(item: any): string[] {
     if (!item) return [];
     const candidates = [item.tags, item.Tags, item.category, item.Category];
     for (const c of candidates) {
       if (!c) continue;
       if (Array.isArray(c)) return c.map(t => String(t).toLowerCase());
       if (typeof c === 'string') return [c.toLowerCase()];
     }
     return [];
   }
 
   function matchesTab(f: any, activeTab: TabKey): boolean {
     const cat = (f.category || f.Category || '').toString().toLowerCase();
     const tags = getTags(f);
     const tagSet = new Set(tags);
     if (activeTab === 'news') return cat === 'news' || tagSet.has('news');
     if (activeTab === 'announcements') return cat.includes('announcement') || tagSet.has('announcements') || tagSet.has('announcement');
     if (activeTab === 'success') return cat === 'success' || tagSet.has('success') || tagSet.has('success story') || tagSet.has('success-stories');
     return true;
   }
 
   const filteredFeatured = useMemo(() => {
      if (activeTab === 'all') return featured;
      return featured.filter(f => matchesTab(f, activeTab));
    }, [featured, activeTab]);
  */
  return (
    <div className={styles.banner}>
      <div
        className={styles.heroSection}
        style={{
          backgroundImage: bannerImages.length > 0
            ? `url('${bannerImages[currentIndex]}')`
            : 'none'
        }}
      >
        <div className={styles.heroOverlay} />
        <div className={styles.heroContent}>
          <div className={styles.centerBlock}>
            <Persona
              text={`Welcome, ${userName}!`}
              size={PersonaSize.size48}
              imageUrl={undefined}          // no image
              hidePersonaDetails={false}    // keep text
              styles={({
                root: { alignItems: 'center', padding: 0 },
                primaryText: {
                  color: '#ffffff',
                  fontSize: '22px',
                  fontWeight: 600,
                  lineHeight: '1.1'
                },
                secondaryText: { display: 'none' },
                coin: { display: 'none' } // hides avatar area (use "coin" slot for Persona)
              } as any)}
            />
          </div>

          <div className={styles.quickLinks}>

            <InfoTile context={context} />
          </div>

          {/* Carousel controls */}
          {bannerImages.length > 1 && (
            <div className={styles.carouselControls}>
              {bannerImages.map((_, idx) => (
                <span
                  key={idx}
                  className={`${styles.dot} ${idx === currentIndex ? styles.activeDot : ''}`}
                  onClick={() => setCurrentIndex(idx)}
                />
              ))}
            </div>
          )}
        </div>
      </div>

      <div className={styles.content}>

        <div className={styles.tabsAndHeader}>
          <div className={styles.tabsContainer}>
            <Pivot
              selectedKey={activeTab}
              onLinkClick={(item) => {
                const key = (item?.props.itemKey || 'all') as TabKey;
                setActiveTab(key);
              }}
              linkFormat="tabs">
              <PivotItem headerText="All" itemKey="all" />
              <PivotItem headerText="News" itemKey="news" />
              <PivotItem headerText="Announcements" itemKey="announcements" />
            </Pivot>
          </div>
        </div>

        <div className={styles.featuredList}>
          {activeTab === 'news' && <NewsList context={context} />}
          {activeTab === 'announcements' && <AnnouncementList context={context} />}
          {activeTab === 'all' && <CombinedFeed context={context} maxItems={6} />}
        </div>
      </div>

      <div className={styles.panelsRow}>
        <div className={styles.sidePanels}>
          {showWelcome && (
            <div className={styles.panel}>
              <WelcomeNewJoiners context={context} />
            </div>
          )}
        </div>

        <div className={styles.fullWidthPanel}>
          <EventsCarousel context={context} excludedSitesCsv={excludedSites} />
        </div>
      </div>


      <div>



      </div>

    </div>
  );
};

export default HeroBanner;