import * as React from 'react';
import { useEffect, useState, useMemo } from 'react';
import {
  Persona,
  PersonaSize,
  IconButton,
  // Stack,
  Pivot,
  PivotItem
} from '@fluentui/react';
import styles from './HeroBanner.module.scss';
import type { IBannerProps } from '../IBannerProps';
import InfoTile from './InfoTile';
import DataService, { IFeaturedItem } from '../../services/DataService';
import NewsList from '../news/NewsList';
import AnnouncementList from '../announcement/AnnouncementList';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

interface IBannerItem {
  Id: number;
  Title: string;
  Active: boolean;
  AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
}

type TabKey = 'all' | 'news' | 'announcements' | 'success';

const HeroBanner: React.FC<IBannerProps> = ({ context }) => {
  const [userName, setUserName] = useState<string>('User');
  const [featured, setFeatured] = useState<IFeaturedItem[]>([]);
  const [activeTab, setActiveTab] = useState<TabKey>('all');
  const [bannerImages, setBannerImages] = useState<string[]>([]);
  const [currentIndex, setCurrentIndex] = useState<number>(0);

  useEffect(() => {
    const svc = new DataService(context);
    svc.getCurrentUserDisplayName().then(name => setUserName(name));
    svc.getFeaturedItems().then(items => setFeatured(items));

    // Fetch banner images from SharePoint list "Banner"
    const url =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Banner')/items` +
      `?$select=Id,Title,Active,AttachmentFiles&$expand=AttachmentFiles`;

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

  function getTags(item: any): string[] {
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
            {/* <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
              {[
                { key: 'helpdesk', label: 'Helpdesk', icon: 'Help' },
                { key: 'directory', label: 'Directory', icon: 'Contact' },
                { key: 'news', label: 'News', icon: 'RSS' },
                { key: 'resources', label: 'Resources', icon: 'OpenFile' },
                { key: 'benefits', label: 'Benefits', icon: 'Money' },
                { key: 'profile', label: 'Profile', icon: 'ContactCard' }
              ].map(link => (
                <InfoTile key={link.key} iconName={link.icon} label={link.label} context={context} />
              ))}
            </Stack> */}
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
              <PivotItem headerText="Success Stories" itemKey="success" />
            </Pivot>
          </div>
        </div>

        <div className={styles.featuredList}>
          {(activeTab === 'news' || activeTab === 'all') && (
            <NewsList context={context} />
          )}

          {(activeTab === 'announcements' || activeTab === 'all') && (
            <AnnouncementList context={context} />
          )}

          {(activeTab === 'all' || activeTab === 'success') && (
            filteredFeatured.length === 0 ? (
              <div className={styles.noItems}>No items to show for this category.</div>
            ) : (
              filteredFeatured.map(item => (
                <div key={item.id} className={styles.featuredCard}>
                  <div className={styles.cardHeader}>
                    <h3 className={styles.cardTitle}>{item.title}</h3>
                  </div>
                  <div className={styles.cardBody}>
                    <p className={styles.cardExcerpt}>{item.excerpt}</p>
                  </div>
                  <div className={styles.cardFooter}>
                    <span className={styles.meta}>{item.views} views</span>
                    <IconButton iconProps={{ iconName: 'Comment' }} title="Comments" ariaLabel="Comments" />
                  </div>
                </div>
              ))
            )
          )}
        </div>
      </div>
    </div>
  );
};

export default HeroBanner;