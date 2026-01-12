import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ICalendarEvent } from '../../services/CalendarEventService';

export type CalendarLayout = 'grid' | 'carousel';
export type CalendarSource = 'outlook' | 'sharepoint' | 'both';

export interface ICalendarProps {
    context: WebPartContext;
    layout: CalendarLayout;
    calendarSource: CalendarSource;
    groupId?: string;
    spSiteUrl?: string;
    spListTitle?: string;
    title?: string;
    isDarkTheme?: boolean;
    itemsPerPage?: number;
}

export interface ICalendarState {
    events: ICalendarEvent[];
    filteredEvents: ICalendarEvent[];
    currentIndex: number;
    loading: boolean;
    selectedDate: Date;
    layout: string;
    error?: string;
}
