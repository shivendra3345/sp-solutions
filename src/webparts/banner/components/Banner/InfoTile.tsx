import * as React from 'react';
import { useEffect, useState } from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './HeroBanner.module.scss';
import { Icon } from '@fluentui/react';

interface IInfoTileItem {
  Id: number;
  Title: string;
  Icon: string;
  Label: string;
  Key: string;
  Active: boolean;
  URL?: string;
  AttachmentFiles?: { FileName: string; ServerRelativeUrl: string }[];
}

interface IInfoTileListProps {
  context?: any;

  onClick?: () => void;
}

const InfoTileList: React.FC<IInfoTileListProps> = ({ context }) => {
  const [tiles, setTiles] = useState<IInfoTileItem[]>([]);

  useEffect(() => {
    const url =
      `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('InfoTiles')/items` +
      `?$select=Id,Title,Icon,Label,Key,Active,URL,AttachmentFiles&$expand=AttachmentFiles`;

    context.spHttpClient.get(url, SPHttpClient.configurations.v1)
      .then((res: SPHttpClientResponse) => res.json())
      .then((data: { value: never[]; }) => {
        const items: IInfoTileItem[] = data.value || [];
        const activeItems = items.filter(i => i.Active);
        setTiles(activeItems);
      })
      .catch((err: any) => console.error("Error fetching InfoTiles:", err));
  }, [context]);

  const getBackgroundImage = (item: IInfoTileItem): string | null => {
    if (item.AttachmentFiles && item.AttachmentFiles.length > 0) {
      return item.AttachmentFiles[0].ServerRelativeUrl;
    }
    return null;
  };

  return (
    <div className={styles.infoTileList}>
      {tiles.map(item => {
        const bgImage = getBackgroundImage(item);
        return (
          <div
            key={item.Key || item.Id}
            className={styles.infoTile}
            style={bgImage ? { backgroundImage: `url(${bgImage})` } : {}}
            onClick={() => {
              if (item.URL) {
                window.open(item.URL, '_blank'); // ✅ open in new tab
              }
            }}
          >
            {!bgImage && (
              <div className={styles.tileCircle}>
                <Icon iconName={item.Icon} className={styles.tileIcon} />
              </div>
            )}
            <div className={styles.tileLabel}>{item.Label}</div>
          </div>
        );
      })}
    </div>
  );
};

export default InfoTileList;