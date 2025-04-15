import { IStoryItem } from '../StoryTileWebPart';

export interface IStoryTileProps {
  description?: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  // For single tile mode
  imageUrl?: string;
  title?: string;
  linkUrl?: string;
  
  // For list mode
  webPartTitle: string;
  storyItems: IStoryItem[];
  isLoading: boolean;
  itemsPerPage: number;
}