import * as React from 'react';
import styles from './StoryTile.module.scss';
import type { IStoryTileProps } from './IStoryTileProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class StoryTile extends React.Component<IStoryTileProps> {
  public render(): React.ReactElement<IStoryTileProps> {
    const {
      description,
      hasTeamsContext,
      title,
      imageUrl,
      linkUrl
    } = this.props;

    const handleTileClick = (): void => {
      window.open(linkUrl, '_blank');
    };

    return (
      <div className={`${styles.storyTile} ${hasTeamsContext ? styles.teams : ''}`}>
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
            <img src={imageUrl} alt={title} />
          </div>
          <div className={styles.contentContainer}>
            <div>
              <h3 className={styles.title}>{escape(title)}</h3>
              <p className={styles.description}>{escape(description)}</p>
            </div>
            <div className={styles.arrowIcon}>
              →
            </div>
          </div>
        </div>
      </div>
    );
  }
}