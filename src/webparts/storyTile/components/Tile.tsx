import * as React from 'react';
import styles from './StoryTile.module.scss';
import { escape } from '@microsoft/sp-lodash-subset';
import { IStoryItem } from '../StoryTileWebPart';
import { Icon } from '@fluentui/react/lib/Icon';

export interface ITileProps {
  item: IStoryItem;
}

const Tile: React.FC<ITileProps> = (props) => {
  const { item } = props;

  const handleTileClick = (): void => {
    window.open(item.linkUrl, '_blank');
  };

  return (
    <div 
      className={styles.tileContainer}
      onClick={handleTileClick}
      role="button"
      tabIndex={0}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          handleTileClick();
        }
      }}
    >
      <div className={styles.imageContainer}>
        <img src={item.imageUrl} alt={item.title} />
      </div>
      <div className={styles.contentContainer}>
        <div>
          <h3 className={styles.title}>{escape(item.title)}</h3>
          <p className={styles.description}>{escape(item.description)}</p>
        </div>
        <div className={styles.arrowIcon}>
          <Icon iconName="ChevronRightMed" />
        </div>
      </div>
    </div>
  );
};

export default Tile;