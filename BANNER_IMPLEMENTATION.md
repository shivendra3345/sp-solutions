# DPT Page Banner Component Implementation

## Overview
A complete banner component system has been created for the DptPage web part with support for two layout types: carousel and tile view. Images are fetched from a SharePoint list, and users can configure the layout and list through the property pane.

## Files Created

### 1. Service Layer
**File:** `src/webparts/dptPage/services/BannerImageService.ts`

- Handles fetching banner images from SharePoint lists
- Interfaces:
  - `IBannerImage`: Defines banner image structure (id, title, imageUrl, description, link)
- Methods:
  - `getBannerImages(listTitle)`: Fetches all banner images from a specified list
  - `getBannerImageById(listTitle, itemId)`: Fetches a single image by ID
- Includes mock data for development/testing
- Handles errors gracefully with fallback to mock data

### 2. Component Layer

**File:** `src/webparts/dptPage/components/PageBanner.tsx`

The main banner component supporting two layouts:

#### Carousel Layout (1-Slide)
- Full-width image carousel
- Previous/Next navigation buttons
- Dot indicators for slide navigation
- Overlay with title, description, and optional link button
- 400px height on desktop, responsive on mobile

#### Tile Layout (1 Large + 2 Small Images)
- Left side: 1 large main image (2:1 grid ratio)
- Right side: 2 smaller images stacked vertically
- Each tile shows title, description, and optional link
- Fully responsive - stacks to single column on mobile

**Features:**
- State management for current slide/images
- Component lifecycle handling (componentDidMount, componentDidUpdate)
- Error handling and loading states
- Responsive design with mobile breakpoints
- Dark theme support

**File:** `src/webparts/dptPage/components/IPageBannerProps.ts`

Interfaces:
- `IPageBannerProps`: Props interface with context, layout type, list title
- `IPageBannerState`: State interface for images, current index, loading, and errors
- `BannerLayout`: Type definition for 'carousel' | 'tile'

**File:** `src/webparts/dptPage/components/PageBanner.module.scss`

Comprehensive styling including:
- Carousel styles (navigation buttons, dots, overlay)
- Tile layout styles (main tile, small tiles column, responsive grid)
- Loading and error states
- Transitions and hover effects
- Dark theme support
- Responsive breakpoints (768px)

### 3. Web Part Configuration

**Updated File:** `src/webparts/dptPage/DptPageWebPart.ts`

Changes:
- Added `bannerLayout` property (carousel | tile)
- Added `bannerListTitle` property (default: 'BannerImages')
- Updated `IDptPageWebPartProps` interface
- Added layout options dropdown
- Updated property pane with:
  - Description field (existing)
  - Layout dropdown selector
  - Banner list title text field

**Updated File:** `src/webparts/dptPage/components/IDptPageProps.ts`

- Added `context` property (WebPartContext)
- Added `bannerLayout` property
- Added `bannerListTitle` property

**Updated File:** `src/webparts/dptPage/components/DptPage.tsx`

- Imported PageBanner component
- Integrated PageBanner at the top of the page
- Passed required props (context, layout, listTitle)

## SharePoint List Requirements

To use this component, you need to create a SharePoint list with the following fields:

| Field Name | Type | Required |
|-----------|------|----------|
| Title | Text | Yes |
| ImageUrl | Text (URL) | Yes |
| Description | Text | No |
| Link | Text (URL) | No |

**Suggested List Name:** `BannerImages`

## Usage

1. **Configure in Property Pane:**
   - Select banner layout (Carousel or Tile View)
   - Enter the SharePoint list title (e.g., "BannerImages")

2. **Populate SharePoint List:**
   - Add items with image URLs, titles, descriptions, and optional links
   - Ensure ImageUrl contains absolute URLs to images

3. **Component renders automatically:**
   - Fetches images on component mount
   - Displays based on selected layout
   - Handles loading and error states

## Responsive Behavior

- **Desktop:** Full-featured with all interactive elements
- **Tablet (768px and below):** 
  - Carousel height reduces to 250px
  - Tile layout adapts to single column
  - Optimized font sizes

## Features Included

✅ Carousel layout with navigation and indicators
✅ Tile layout with responsive grid
✅ SharePoint list integration
✅ Error handling and mock data fallback
✅ Property pane customization
✅ Loading states
✅ Dark theme support
✅ Fully responsive design
✅ Accessibility features (aria-labels, keyboard navigation)
✅ Smooth transitions and animations

## Next Steps

1. Create a SharePoint list named "BannerImages" (or your preferred name)
2. Add the required fields (Title, ImageUrl, Description, Link)
3. Populate with banner data
4. Deploy the web part
5. Configure the layout and list title in the property pane
