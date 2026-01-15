# Team Members Component — Implementation Notes

This document explains how the `TeamMembers` component is implemented, how it fetches data from SharePoint, the expected list schema, styling, props, and how to customize its behavior.

Location
- Component: `src/webparts/dptPage/components/TeamMembers/TeamMembers.tsx`
- SCSS: `src/webparts/dptPage/components/TeamMembers/TeamMembers.module.scss`
- Props/interface: `src/webparts/dptPage/components/TeamMembers/ITeamMembersProps.ts` (component props) 
- Service: `src/webparts/dptPage/services/TeamMembersService.ts`

Summary
- Renders a responsive grid of team member cards (photo, name, job title, contact, location).
- Data is fetched from a SharePoint list (default list title: `Team Members`) via `TeamMembersService` using the REST API.
- Supports loading, error and empty states.
- Supports a dark theme via `isDarkTheme` prop.
- Supports showing/hiding from the web part property pane via `showTeamMembers`.

Props (passed from web part)
- `context: WebPartContext` — SPFx web part context used by the service.
- `listTitle?: string` — SharePoint list title to read members from. Default: `Team Members`.
- `title?: string` — Header/title string displayed above the grid.
- `itemsPerRow?: number` — optional control for number of cards per row (component currently uses CSS grid so this is advisory / not strictly required).
- `isDarkTheme?: boolean` — toggle dark theme styles.

Service: TeamMembersService
- Location: `src/webparts/dptPage/services/TeamMembersService.ts`
- Exports interface `ITeamMember`:
  - `id: string`
  - `displayName: string` (maps to `Title` field)
  - `jobTitle?: string` (maps to `JobTitle`)
  - `email?: string` (maps to `Email`)
  - `phone?: string` (maps to `Phone`)
  - `location?: string` (maps to `Location`)
  - `photoUrl?: string` (maps to `PhotoUrl`)
- Methods:
  - `getTeamMembers(listTitle: string): Promise<ITeamMember[]>` — fetches up to top 100 items ordered by Title
  - `getTeamMemberById(listTitle: string, itemId: number): Promise<ITeamMember | null>`
  - `getDefaultAvatar(displayName: string): string` — returns a placeholder avatar URL if photo missing
- Important: The service falls back to mock data if `context` or `spHttpClient` is not available (useful for local development).

Expected SharePoint List Schema (recommended columns)
- Column internal names / display names expected by the service (case-sensitive in REST `$select`):
  - `Title` (Single line of text) — used for `displayName`
  - `JobTitle` (Single line of text)
  - `Email` (Single line of text)
  - `Phone` (Single line of text)
  - `Location` (Single line of text)
  - `PhotoUrl` (Single line of text) — preferably the absolute URL to the user/photo image

If your list uses different column names, update `TeamMembersService.mapItemToTeamMember` to map the correct fields.

Rendering and Markup
- The component renders:
  - Wrapper: `div` with `styles.pageTeamMembers`
  - Header (if `title` is provided): `div` with `styles.membersHeader` and `h2` `styles.membersTitle`
  - States:
    - Loading: `styles.loading` block with spinner
    - Error: `styles.error` with retry button
    - Empty: `styles.noMembers` message
  - Grid: `div` with `styles.membersGrid` (CSS Grid with `minmax` handling)
  - Card: `div` with `styles.memberCard` containing:
    - Photo: `div` `styles.memberPhoto` + `img` (or placeholder)
    - Info: `div` `styles.memberInfo` with name, job, contact, location

Styling
- SCSS file: `TeamMembers.module.scss` contains variables for colors and spacing. Recent updates made the header style consistent with the Calendar and QuickLinks components.
- Grid: Uses `grid-template-columns: repeat(auto-fill, minmax(Xpx, 1fr))` for responsiveness.
- Card: compact design with subtle shadow, reduced padding and smaller photo height (to fit more cards per row).

Accessibility
- Each card has `role="button"` and `tabIndex=0` where clickable behavior is required (if you add click handlers to cards, ensure keyboard handling for Enter/Space is implemented).
- Email and phone links use `mailto:` and `tel:` anchors.
- Images should include `alt` text (currently set to the member's display name).

Theming
- The component uses a `isDarkTheme` prop to add a `.darkTheme` modifier class and apply dark background / lighter text colors.
- To support a site-level theme, forward the theme value from the web part into the component.

Customization
- Field mapping: if your list columns differ, change `mapItemToTeamMember` in `TeamMembersService.ts`.
- Photo handling: if photo is stored in an `Image` column (complex object), update the REST `$select` and mapping accordingly to extract the URL.
- Pagination: currently the service fetches up to 100 items; for larger lists, implement paging in the service and a "Load more" in the UI.

How to change the grid density
- Update SCSS `membersGrid` `minmax(...)` values to increase/decrease the minimum card width.
- Alternatively, add a prop `itemsPerRow` and compute `grid-template-columns` in inline styles, but keep accessibility and responsiveness in mind.

Testing
- The service provides mock data so you can test the component without SP context.
- Manual checks:
  - Add a SharePoint list matching the schema and set the web part `Team Members List Title` property.
  - Toggle `Show Team Members` from the property pane to ensure show/hide works.
  - Verify dark theme styles by toggling `isDarkTheme` in the web part properties (or pass site theme value).

Common issues and fixes
- Missing photos: TeamMembersService provides `getDefaultAvatar` to generate a placeholder. Ensure `photoUrl` is empty or invalid for placeholder.
- Different field names: update mapping in `TeamMembersService.mapItemToTeamMember`.
- Very large lists: implement server-side pagination in `getTeamMembers` using `$skiptoken` or indexed filtered queries.

Developer notes
- File paths referenced in codebase:
  - `src/webparts/dptPage/components/TeamMembers/TeamMembers.tsx`
  - `src/webparts/dptPage/components/TeamMembers/TeamMembers.module.scss`
  - `src/webparts/dptPage/components/TeamMembers/ITeamMembersProps.ts`
  - `src/webparts/dptPage/services/TeamMembersService.ts`
- The web part passes props from `DptPageWebPart.ts` → `DptPage.tsx` → `TeamMembers`.

Change log (recent)
- Header and title sizes were reduced to match other components (Calendar, QuickLinks).
- Card dimensions, photo heights, padding and typography were made more compact.
- QuickLinks was added and the page layout changed to ensure alignment with Team Members and Calendar.

If you'd like, I can:
- Create a small sample SharePoint list JSON template you can import (site columns & sample items).
- Add `itemsPerRow` prop support to control grid layout from the web part properties.
- Implement lazy loading for member photos for better performance.

---

If you want a separate file placed elsewhere or additional details (list JSON, field internal names, or example mock data), tell me where and I will add it.