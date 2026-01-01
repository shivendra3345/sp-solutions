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

const HeroBanner: React.FC<IBannerProps> = ({ context, showWelcome, excludedSites, showNews = true, showEvents = true, welcomeTitle, eventsTitle, newsTitle, selectedListIds, descriptionFieldNames }) => {
  const [userName, setUserName] = useState<string>('User');
  //const [featured, setFeatured] = useState<IFeaturedItem[]>([]);
  const [activeTab, setActiveTab] = useState<TabKey>('all');

  const containerRef = React.useRef<HTMLDivElement | null>(null);
  const [minHeight, setMinHeight] = useState<number | undefined>(undefined);
  const newsRef = React.useRef<HTMLDivElement | null>(null);
  const annRef = React.useRef<HTMLDivElement | null>(null);
  const allRef = React.useRef<HTMLDivElement | null>(null);

  const measureHeights = React.useCallback(() => {
    try {
      const h1 = newsRef.current ? newsRef.current.offsetHeight : 0;
      const h2 = annRef.current ? annRef.current.offsetHeight : 0;
      const h3 = allRef.current ? allRef.current.offsetHeight : 0;
      const maxh = Math.max(h1, h2, h3, 220);
      setMinHeight(maxh);
    } catch (e) {
      setMinHeight(320);
    }
  }, []);
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
  // animate tab switches: fade out current, swap content, fade in
  React.useEffect(() => {
    // When the active tab changes, measure heights to avoid layout jumps,
    // then clear the temporary minHeight after animations settle.
    measureHeights();
    let t1: number | undefined;
    let t2: number | undefined;
    t1 = window.setTimeout(() => {
      // allow incoming content to render and then clear minHeight
      t2 = window.setTimeout(() => {
        setMinHeight(undefined);
      }, 360);
    }, 160);
    return () => {
      if (t1) clearTimeout(t1);
      if (t2) clearTimeout(t2);
    };
  }, [activeTab, measureHeights]);

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

      {showNews && (
        <div className={styles.content}>

          {newsTitle && (<h3 className={styles.title}>{newsTitle}</h3>)}

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

          <div className={styles.tabContainer} ref={containerRef} style={minHeight ? { minHeight: `${minHeight}px` } : undefined}>
            {/* Always render all layers (hidden/visible) and measure heights to avoid layout jumps */}
            <div ref={newsRef} className={`${styles.tabLayer} ${activeTab === 'news' ? styles.visible : styles.hidden}`} aria-hidden={activeTab !== 'news'}>
              <div className={styles.featuredList}><NewsList context={context} /></div>
            </div>
            <div ref={annRef} className={`${styles.tabLayer} ${activeTab === 'announcements' ? styles.visible : styles.hidden}`} aria-hidden={activeTab !== 'announcements'}>
              <div className={styles.featuredList}><AnnouncementList context={context} /></div>
            </div>
            <div ref={allRef} className={`${styles.tabLayer} ${activeTab === 'all' ? styles.visible : styles.hidden}`} aria-hidden={activeTab !== 'all'}>
              <div className={styles.featuredList}><CombinedFeed context={context} maxItems={6} /></div>
            </div>
          </div>

        </div>
      )}

      <div className={styles.panelsRow}>
        {showWelcome && (
          <div className={styles.sidePanels}>
            <div className={styles.panel}>
              <WelcomeNewJoiners context={context} title={welcomeTitle} />
            </div>
          </div>
        )}

        {showEvents && (
          <div className={styles.fullWidthPanel}>
            <EventsCarousel context={context} excludedSitesCsv={excludedSites} selectedListIdsCsv={selectedListIds} title={eventsTitle} descriptionFieldNamesCsv={descriptionFieldNames} />
          </div>
        )}
      </div>


      <div>



      </div>

    </div >
  );
};

export default HeroBanner;