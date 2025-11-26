import * as React from 'react';
import { Icon } from '@fluentui/react';
import styles from './HeroBanner.module.scss';

export interface IInfoTileProps {
  iconName: string;
  label: string;
  onClick?: () => void;
}

const InfoTile: React.FC<IInfoTileProps> = ({ iconName, label, onClick }) => {
  return (
    <div className={styles.infoTile} onClick={onClick}>
      <div className={styles.tileCircle}>
        <Icon iconName={iconName} className={styles.tileIcon} />
      </div>
      <div className={styles.tileLabel}>{label}</div>
    </div>
  );
};

export default InfoTile;