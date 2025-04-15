export interface IStoryTileProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  
  // New properties for the story tile
  imageUrl: string;
  title: string;
  linkUrl: string;
}