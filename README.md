# Story Tile Web Part for SharePoint

A modern SharePoint Framework (SPFx) web part that displays content from a SharePoint list in a tile-based layout with carousel functionality.

## Features

- Displays 1-4 tiles per view with configurable layout
- Carousel navigation for browsing through additional tiles
- Pulls content dynamically from a SharePoint list using PnP JS
- Robust handling of SharePoint Image fields
- Responsive design that works across all device sizes
- Modern UI with shadows, rounded corners, and hover effects
- Configurable fields for customization

![Story Tile Web Part](./assets/story-tile-preview.png)

## Getting Started

### Prerequisites

- Node.js (version 18.17.1 or higher)
- SharePoint Developer environment
- SharePoint list with the correct content type (see below)

### Installation

1. Clone this repository
2. Run `npm install`
3. Run `gulp serve` to test locally
4. Run `gulp bundle --ship` and `gulp package-solution --ship` to package for deployment
5. Upload the `.sppkg` file from the `sharepoint/solution` folder to your SharePoint App Catalog
6. Add the web part to your page

## SharePoint List Setup

### Content Type Columns

Create a SharePoint list with the following columns:

1. **Title** (Default column)
   - Used for the tile heading

2. **Description**
   - Type: Multiple lines of text
   - Used for the brief description on the tile

3. **Image**
   - Type: Image
   - The image to display on the tile
   - The web part has robust handling for this field type

4. **LinkURL**
   - Type: Single line of text
   - The URL to navigate to when the tile is clicked

5. **SortOrder** (Optional)
   - Type: Number
   - Used to control the order of tiles

### Creating the List

1. Create a new list in SharePoint
2. Add the columns specified above
3. Add your content items
4. Configure the web part to use this list

## Web Part Configuration

In the web part properties pane, you can configure:

- **Web Part Title**: The heading displayed above the tiles
- **SharePoint List Name**: Name of the list containing your tile content
- **Tiles Per View**: Number of tiles to display at once (1-4)
- **Field Name Settings**: Configure custom field names if they differ from defaults

## Technical Details

### PnP JS Integration

This web part uses PnP JS (SharePoint Patterns and Practices JavaScript library) for SharePoint data operations, which offers several advantages:

- Cleaner, more maintainable code for SharePoint operations
- Improved error handling and fallback mechanisms
- Better performance through optimized queries
- More robust handling of various SharePoint field types

### Image Field Handling

The web part implements a multi-layered approach to retrieve image data from SharePoint:

1. **Field Detection**: Automatically detects the internal name of the Image field
2. **Multiple Retrieval Methods**:
   - Uses `fieldValuesAsHtml` to extract image URLs from HTML representation
   - Attempts direct access to image field data in various formats
   - Uses RenderListDataAsStream as a fallback method
3. **URL Processing**: Handles both relative and absolute URLs correctly

This approach ensures maximum compatibility with different SharePoint configurations and field types.

## Development Notes

- Built using SharePoint Framework (SPFx) 1.20.0
- Uses React and Fluent UI components
- Implements responsive grid layout with CSS Grid
- Leverages PnP JS library for enhanced SharePoint operations

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## License

This project is licensed under the MIT License - see the LICENSE file for details.
