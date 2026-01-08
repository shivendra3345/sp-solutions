# DPT Page Calendar Component Implementation

## Overview
A complete calendar component system has been created for the DptPage web part with support for fetching events from both Outlook Group Calendars and SharePoint Calendar lists. The component supports two layout types: Calendar Grid View and Carousel View.

## Features

### Event Fetching
- **Outlook Group Calendar**: Fetch events from Microsoft 365 Group calendars using Microsoft Graph API
- **SharePoint Calendar**: Fetch events from SharePoint calendar lists
- **Dual Source**: Combine events from both Outlook and SharePoint
- **Event Fields**: Title, Description, Start Date, End Date, Created By, Location, Organizer, Thumbnail Image

### Layout Options

#### Calendar Grid View
- Monthly calendar display
- Shows all events for each day
- Day cells highlight today's date
- Navigation buttons for previous/next month
- Compact event display with "more events" indicators
- Responsive design

#### Carousel View
- Full-featured event display with thumbnail images
- Shows one event at a time
- Previous/Next navigation buttons
- Dot indicators for event navigation
- Displays complete event details (title, time, description, location, organizer)
- Default placeholder images if no thumbnail available

## Files Created

### Service Layer

**File:** `src/webparts/dptPage/services/CalendarEventService.ts`

Interfaces:
- `ICalendarEvent`: Defines event structure with all event properties

Methods:
- `getGroupCalendarEvents(groupId, startDate, endDate)`: Fetch events from Outlook Group Calendar
- `getSharePointCalendarEvents(siteUrl, listTitle, startDate, endDate)`: Fetch events from SharePoint calendar list
- `getMyCalendarEvents(startDate, endDate)`: Fetch current user's calendar events
- `getDefaultThumbnail()`: Generate random placeholder images

Features:
- Automatic date filtering
- Error handling with mock data fallback
- Mock calendar data for development/testing
- Proper date/time formatting

### Component Layer

**File:** `src/webparts/dptPage/components/PageCalendar.tsx`

Main calendar component supporting two layouts:

#### Calendar Grid View Features
- Monthly view with day-by-day event display
- Today's date highlighting
- Month navigation
- Event preview in calendar cells
- Responsive grid layout

#### Carousel View Features
- Large thumbnail images (with fallback to generated placeholders)
- Complete event information display
- Event time and date formatting
- Location and organizer information
- Navigation controls and indicators
- Event counter (e.g., "Event 1 of 5")

**Component Methods:**
- `componentDidMount()`: Load events on initial load
- `componentDidUpdate()`: Reload when props change
- `loadCalendarEvents()`: Fetch events from configured sources
- `formatEventTime()`: Format event time display
- `formatDateRange()`: Format date range display
- Layout-specific rendering methods

**File:** `src/webparts/dptPage/components/IPageCalendarProps.ts`

Type Definitions:
- `CalendarLayout`: 'grid' | 'carousel'
- `CalendarSource`: 'outlook' | 'sharepoint' | 'both'

Interfaces:
- `IPageCalendarProps`: Component props with all configuration options
- `IPageCalendarState`: Component state management

**File:** `src/webparts/dptPage/components/PageCalendar.module.scss`

Comprehensive styling including:
- Header and navigation styles
- Calendar grid layout with responsive design
- Carousel container and controls
- Event card styling
- Loading and error states
- Dark theme support
- Responsive breakpoints (768px)
- Smooth transitions and animations

### Web Part Configuration

**Updated File:** `src/webparts/dptPage/DptPageWebPart.ts`

New Properties:
- `calendarLayout`: 'grid' | 'carousel'
- `calendarSource`: 'outlook' | 'sharepoint' | 'both'
- `groupId`: Microsoft 365 Group ID for Outlook calendar
- `spSiteUrl`: SharePoint site URL
- `spListTitle`: SharePoint calendar list title

Property Pane Groups:
1. **Banner Settings** - Banner-related configurations
2. **Calendar Settings** - Calendar-related configurations

**Updated File:** `src/webparts/dptPage/components/IDptPageProps.ts`

- Added `calendarLayout` property
- Added `calendarSource` property
- Added `groupId`, `spSiteUrl`, `spListTitle` properties

**Updated File:** `src/webparts/dptPage/components/DptPage.tsx`

- Integrated PageCalendar component below PageBanner
- Passes all required props from web part to calendar component

## SharePoint Setup Requirements

### For SharePoint Calendar Events
Create a SharePoint calendar list with the following fields:

| Field Name | Type | Required |
|-----------|------|----------|
| Title | Text | Yes |
| Event Date (EventDate) | Date/Time | Yes |
| End Date (EndDate) | Date/Time | Yes |
| Description | Multiple lines of text | No |
| Location | Text | No |
| Created | System Field | Auto |
| Created By | System Field | Auto |

**Suggested List Name:** `Calendar`

### For Outlook Group Calendar
- Ensure you have access to the Microsoft 365 Group calendar
- Get the Group ID from Microsoft Teams admin center or Graph Explorer
- The component uses Microsoft Graph API to fetch events

## Property Pane Configuration

### Banner Settings
- **Description**: Web part description text
- **Banner Layout**: Choose between Carousel (1-Slide) or Tile View
- **Banner List Title**: Name of the SharePoint list containing banner images

### Calendar Settings
- **Calendar Layout**: Choose between Calendar Grid View or Carousel View
- **Calendar Source**: Select Outlook, SharePoint, or Both
- **Outlook Group Calendar ID**: Enter the ID of the Microsoft 365 Group (required for Outlook source)
- **SharePoint Calendar List Title**: Enter the name of the calendar list (default: "Calendar")

## Event Display

### Calendar Grid View
- Shows month view with all events for each day
- Today's date is highlighted in blue
- Events displayed as colored badges within day cells
- "More events" indicator when day has more than 3 events
- Click month navigation to view different months

### Carousel View
- Displays full event details with large thumbnail image
- Event title, date range, and time
- Location, organizer, and description information
- Navigation buttons to browse through events
- Dot indicators show current position
- Event counter displays progress (e.g., "Event 1 of 5")

## Mock Data

For development and testing, the service includes mock calendar events with:
- Team Standup meeting
- Project Review meeting
- All Hands Meeting
- Training Workshop
- Client Presentation

Mock data is automatically used when:
- No group ID is provided
- No SharePoint list is configured
- API calls fail

## Responsive Behavior

- **Desktop**: Full-featured with all interactive elements
- **Tablet (768px and below)**: 
  - Calendar grid cells reduce height
  - Carousel layout adapts with stacked image and content
  - Optimized font sizes
  - Single-column layout where applicable

## Features Included

✅ Fetch events from Outlook Group Calendar
✅ Fetch events from SharePoint Calendar List
✅ Combine events from both sources
✅ Calendar Grid View with month navigation
✅ Carousel View with full event details
✅ Thumbnail images with random placeholders
✅ Event filtering by date range
✅ Error handling with graceful fallback
✅ Loading states
✅ Dark theme support
✅ Fully responsive design
✅ Accessibility features (aria-labels)
✅ Smooth transitions and animations
✅ Property pane customization
✅ Mock data for development

## Usage Examples

### Configuration 1: Outlook Calendar Only
1. Set **Calendar Source** to "Outlook Group Calendar"
2. Enter the **Outlook Group Calendar ID**
3. Choose layout preference
4. Events from the group calendar will be displayed

### Configuration 2: SharePoint Calendar Only
1. Set **Calendar Source** to "SharePoint Calendar"
2. Enter the **SharePoint Calendar List Title**
3. Ensure the SharePoint calendar list has required fields
4. Choose layout preference

### Configuration 3: Combined View
1. Set **Calendar Source** to "Both Outlook & SharePoint"
2. Configure both Group Calendar ID and SharePoint list title
3. Events from both sources will be merged and sorted by date
4. Choose layout preference

## API Integration

### Microsoft Graph API
- Endpoint: `/groups/{groupId}/calendar/calendarview`
- Authentication: Azure AD with appropriate permissions
- Required Permissions: `Calendars.Read`

### SharePoint REST API
- Endpoint: `/_api/web/lists/getbytitle('{listTitle}')/items`
- Authentication: User's current context
- Filtering: By EventDate and EndDate

## Styling

The component includes:
- Default: Light theme with clean, modern design
- Dark Theme: Automatically applied when enabled
- Responsive Grid System: Adapts to screen size
- Smooth Animations: Transitions on interactions
- Accessibility Colors: High contrast text and buttons

## Next Steps

1. Create a SharePoint Calendar list with required fields
2. (Optional) Configure Outlook Group Calendar access
3. Deploy the web part
4. Configure the layout and data sources in the property pane
5. Customize colors and styling in SCSS if needed
