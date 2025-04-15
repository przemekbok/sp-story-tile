import * as React from 'react';
import styles from './StoryTile.module.scss';
import { IStoryTileProps } from './IStoryTileProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Tile from './Tile';
import { IStoryItem } from '../StoryTileWebPart';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Icon } from '@fluentui/react/lib/Icon';

export default class StoryTile extends React.Component<IStoryTileProps, {
  currentPage: number;
}> {
  
  constructor(props: IStoryTileProps) {
    super(props);
    this.state = {
      currentPage: 0
    };
  }

  public render(): React.ReactElement<IStoryTileProps> {
    const { 
      webPartTitle, 
      storyItems, 
      isLoading, 
      itemsPerPage,
      hasTeamsContext 
    } = this.props;

    const { currentPage } = this.state;
    
    // Calculate total pages
    const totalPages = Math.ceil(storyItems.length / itemsPerPage);
    
    // Get items for current page
    const startIndex = currentPage * itemsPerPage;
    const endIndex = startIndex + itemsPerPage;
    const currentItems = storyItems.slice(startIndex, endIndex);
    
    // Generate grid class based on items per page
    let gridClass = styles.gridOne;
    if (itemsPerPage === 2) {
      gridClass = styles.gridTwo;
    } else if (itemsPerPage === 3) {
      gridClass = styles.gridThree;
    } else if (itemsPerPage === 4) {
      gridClass = styles.gridFour;
    }

    // Handle navigation
    const goToPreviousPage = (): void => {
      if (currentPage > 0) {
        this.setState({ currentPage: currentPage - 1 });
      }
    };

    const goToNextPage = (): void => {
      if (currentPage < totalPages - 1) {
        this.setState({ currentPage: currentPage + 1 });
      }
    };

    return (
      <div className={`${styles.storyTile} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.container}>
          <h2 className={styles.webPartTitle}>{escape(webPartTitle)}</h2>
          
          {isLoading ? (
            <div className={styles.spinner}>
              <Spinner size={SpinnerSize.large} label="Loading items..." />
            </div>
          ) : storyItems.length === 0 ? (
            <div className={styles.noItems}>
              <p>No items found. Please check the list configuration in the web part properties.</p>
            </div>
          ) : (
            <div className={styles.carouselContainer}>
              <div className={`${styles.tilesGrid} ${gridClass}`}>
                {currentItems.map((item: IStoryItem) => (
                  <Tile 
                    key={item.id} 
                    item={item} 
                  />
                ))}
              </div>
              
              {totalPages > 1 && (
                <div className={styles.navigationControls}>
                  <button 
                    className={`${styles.navButton} ${currentPage === 0 ? styles.disabled : ''}`}
                    onClick={goToPreviousPage}
                    disabled={currentPage === 0}
                    aria-label="Previous page"
                  >
                    <Icon iconName="ChevronLeftMed" />
                  </button>
                  <div className={styles.pageIndicator}>
                    {`${currentPage + 1} / ${totalPages}`}
                  </div>
                  <button 
                    className={`${styles.navButton} ${currentPage === totalPages - 1 ? styles.disabled : ''}`}
                    onClick={goToNextPage}
                    disabled={currentPage === totalPages - 1}
                    aria-label="Next page"
                  >
                    <Icon iconName="ChevronRightMed" />
                  </button>
                </div>
              )}
            </div>
          )}
        </div>
      </div>
    );
  }
}