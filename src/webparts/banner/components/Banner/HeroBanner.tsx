import * as React from 'react';
import { useEffect, useState, useMemo } from 'react';
import {
  SearchBox,
  Persona,
  PersonaSize,
  IconButton,
  Stack,
  Pivot,
  PivotItem
} from '@fluentui/react';
import styles from './HeroBanner.module.scss';
import type { IBannerProps } from '../IBannerProps';
import InfoTile from './InfoTile';
import DataService, { IFeaturedItem } from '../../services/DataService';

const quickLinks = [
  { key: 'helpdesk', label: 'Helpdesk', icon: 'Help' },
  { key: 'directory', label: 'Directory', icon: 'Contact' },
  { key: 'news', label: 'News', icon: 'RSS' },
  { key: 'resources', label: 'Resources', icon: 'OpenFile' },
  { key: 'benefits', label: 'Benefits', icon: 'Money' },
  { key: 'profile', label: 'Profile', icon: 'ContactCard' }
];

type TabKey = 'all' | 'news' | 'announcements' | 'success';

const HeroBanner: React.FC<IBannerProps & { heroImageUrl?: string }> = ({ context, heroImageUrl }) => {
  const [userName, setUserName] = useState<string>('User');
  const [featured, setFeatured] = useState<IFeaturedItem[]>([]);
  const [activeTab, setActiveTab] = useState<TabKey>('all');

  useEffect(() => {
    const svc = new DataService(context);
    svc.getCurrentUserDisplayName().then(name => setUserName(name));
    svc.getFeaturedItems().then(items => setFeatured(items));
  }, [context]);

  // Helper: normalize tags for an item into a lowercase string array
function getTags(item: any): string[] {
  if (!item) return [];

  // Common shapes to check (adjust to your data)
  const candidates = [
    item.tags,            // lowercase property used earlier
    item.Tags,            // PascalCase
    item.tagsString,      // comma-separated string
    item.TagsString,
    item.Category ? [item.category, item.Category] : undefined, // sometimes category is used
    item.categories,      // plural
    item.taxonomy,        // managed metadata object/array
    item.ManagedMetadata, // other naming
    item.customTags       // custom field
  ];

  // If any candidate is an array, flatten and return strings
  for (const c of candidates) {
    if (!c) continue;
    if (Array.isArray(c)) {
      return c.map(t => String(t).toLowerCase());
    }
    if (typeof c === 'string') {
      // comma or semicolon separated
      return c.split(/[,;|]/).map(s => s.trim().toLowerCase()).filter(Boolean);
    }
    // If it's an object (e.g., taxonomy), try to extract label fields
    if (typeof c === 'object') {
      // handle array-like objects (e.g., [{Label:'x'}]) or single object
      if (Array.isArray(c)) {
        const out: string[] = [];
        c.forEach((entry: any) => {
          if (!entry) return;
          if (typeof entry === 'string') out.push(entry.toLowerCase());
          else if (entry.Label) out.push(String(entry.Label).toLowerCase());
          else if (entry.Name) out.push(String(entry.Name).toLowerCase());
          else if (entry.title) out.push(String(entry.title).toLowerCase());
        });
        if (out.length) return out;
      } else {
        // single object
        if (c.Label) return [String(c.Label).toLowerCase()];
        if (c.Name) return [String(c.Name).toLowerCase()];
        if (c.title) return [String(c.title).toLowerCase()];
      }
    }
  }

  // fallback: try to read a generic field that might contain tags
  if (item.category) return [String(item.category).toLowerCase()];
  if (item.Category) return [String(item.Category).toLowerCase()];

  return [];
}
  const heroStyle: React.CSSProperties = heroImageUrl
    ? ({ ['--hero-bg' as any]: `url('${heroImageUrl}')` } as React.CSSProperties)
    : {};

  // Filter featured items based on active tab.
  // Adjust the property names (category, tags) to match your IFeaturedItem shape if needed.

// helper: normalize tags for an item (assumes getTags exists and returns lowercase strings)
function matchesTab(f: any, activeTab: 'all' | 'news' | 'announcements' | 'success'): boolean {
  const cat = (f.category || f.Category || '').toString().toLowerCase();
  const tags = getTags(f); // getTags should return an array of lowercase strings
  const tagSet = new Set(tags.map(t => t.toLowerCase()));

  const hasTag = (key: string) => tagSet.has(key.toLowerCase());
  const hasAnyTag = (...keys: string[]) => keys.some(k => hasTag(k));

  if (activeTab === 'news') {
    return cat === 'news' || hasTag('news');
  }

  if (activeTab === 'announcements') {
    return (cat === 'announcement' || cat === 'announcements') || hasAnyTag('announcement', 'announcements');
  }

  if (activeTab === 'success') {
    return cat === 'success' || hasAnyTag('success', 'success story', 'success-stories');
  }

  return true;
}// updated useMemo that uses matchesTab
const filteredFeatured = useMemo(() => {
  if (activeTab === 'all') return featured;
  return featured.filter(f => matchesTab(f, activeTab));
}, [featured, activeTab]);


  return (
    <div className={styles.banner}>
      <div className={styles.heroSection} style={heroStyle}>
        <div className={styles.heroOverlay} />
        <div className={styles.heroContent}>
          <div className={styles.centerBlock}>
            <Persona
              text={`Welcome, ${userName}!`}
              size={PersonaSize.size48}
              styles={{
                root: { alignItems: 'center', padding: 0 },
               // image: { display: 'none' },
                primaryText: {
                  color: '#ffffff',
                  fontSize: '22px',
                  fontWeight: 600,
                  lineHeight: '1.1'
                },
                secondaryText: { display: 'none' }
              }}
            />

            <div className={styles.search}>
              <SearchBox placeholder="Search Topics & Templates" />
            </div>
          </div>

          <div className={styles.quickLinks}>
            <Stack horizontal wrap tokens={{ childrenGap: 12 }}>
              {quickLinks.map(link => (
                <InfoTile key={link.key} iconName={link.icon} label={link.label} />
              ))}
            </Stack>
          </div>
        </div>
      </div>

      {/* Tabs + Section header (placed before the featured list) */}
      <div className={styles.content}>
        <div className={styles.tabsAndHeader}>
          <div className={styles.tabsContainer}>
            <Pivot
              selectedKey={activeTab}
              onLinkClick={(item) => {
                const key = (item?.props.itemKey || 'all') as TabKey;
                setActiveTab(key);
              }}
              linkFormat="tabs"
            >
              <PivotItem headerText="All" itemKey="all" />
              <PivotItem headerText="News" itemKey="news" />
              <PivotItem headerText="Announcements" itemKey="announcements" />
              <PivotItem headerText="Success Stories" itemKey="success" />
            </Pivot>
          </div>

           
        </div>

        {/* Featured list filtered by tab */}
        <div className={styles.featuredList}>
          {filteredFeatured.length === 0 ? (
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
          )}
        </div>

        <div className={styles.sidePanels}>
          <div className={styles.panel}>
            <h4>STAFF DIRECTORY</h4>
            <div className={styles.staffList}>
              <Persona text="Andres Clark" secondaryText="Product Designer" size={PersonaSize.size40} />
              <Persona text="Andrew Colton" secondaryText="Software Engineer" size={PersonaSize.size40} />
              <Persona text="Emily Hunt" secondaryText="HR Manager" size={PersonaSize.size40} />
            </div>
          </div>

          <div className={styles.panel}>
            <h4>CALENDAR</h4>
            <div className={styles.calendar}>
              <div className={styles.calendarMonth}>October 2023</div>
              <ul className={styles.events}>
                <li>10:00 AM - Team Standup</li>
                <li>11:00 AM - HR Meeting</li>
                <li>01:00 PM - Product Sync</li>
              </ul>
            </div>
          </div>
        </div>

        <div className={styles.metrics}>
          <div className={styles.metric}>Latest Growth <strong>5.3%</strong></div>
          <div className={styles.metric}>Employee Growth <strong>5.2%</strong></div>
          <div className={styles.metric}>Annual Access <strong>27.6K</strong></div>
          <div className={styles.metric}>Employee Reach <strong>15.3M</strong></div>
        </div>
      </div>
    </div>
  );
};

export default HeroBanner;