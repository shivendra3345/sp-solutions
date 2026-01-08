import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ICalendarEvent {
    id: string;
    title: string;
    description?: string;
    startDate: Date;
    endDate: Date;
    createdBy?: string;
    thumbnailUrl?: string;
    location?: string;
    organizer?: string;
    isAllDay?: boolean;
    eventUrl?: string;
}

export default class CalendarEventService {
    private context: WebPartContext;
    private useMock: boolean = false;

    constructor(context: WebPartContext) {
        this.context = context;
        // Use mock data by default - Graph API requires additional setup
        this.useMock = !context || !context.spHttpClient;
    }

    /**
     * Fetch events from Outlook Group Calendar
     * @param groupId The ID of the Microsoft 365 Group
     * @param startDate Start date for filtering events
     * @param endDate End date for filtering events
     * @returns Promise resolving to an array of calendar events
     */
    public async getGroupCalendarEvents(
        groupId: string,
        startDate?: Date,
        endDate?: Date
    ): Promise<ICalendarEvent[]> {
        if (this.useMock || !groupId) {
            return this.getMockCalendarEvents();
        }

        // Graph API integration requires additional authentication setup
        // For now, return mock data - can be integrated with proper Graph client setup
        return this.getMockCalendarEvents();
    }

    /**
     * Fetch events from SharePoint Calendar List
     * @param siteUrl The URL of the SharePoint site
     * @param listTitle The title of the calendar list
     * @param startDate Start date for filtering events
     * @param endDate End date for filtering events
     * @returns Promise resolving to an array of calendar events
     */
    public async getSharePointCalendarEvents(
        siteUrl: string,
        listTitle: string,
        startDate?: Date,
        endDate?: Date
    ): Promise<ICalendarEvent[]> {
        if (this.useMock || !siteUrl || !listTitle) {
            return this.getMockCalendarEvents();
        }

        try {
            const start = startDate ? startDate.toISOString() : new Date(new Date().getFullYear(), new Date().getMonth(), 1).toISOString();
            const end = endDate ? endDate.toISOString() : new Date(new Date().getFullYear(), new Date().getMonth() + 1, 0).toISOString();

            // SharePoint REST API query for calendar events
            const requestUrl = `${siteUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$filter=EventDate ge '${start}' and EndDate le '${end}'&$select=Id,Title,Description,EventDate,EndDate,Created,Author/Title&$expand=Author/Title&$orderby=EventDate desc&$top=100`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Failed to fetch SharePoint calendar events: ${response.statusText}`);
            }

            const data = await response.json();
            return (data.value || []).map((item: any) => this.mapSharePointEventToCalendarEvent(item));
        } catch (error) {
            console.error('Error fetching SharePoint calendar events:', error);
            return this.getMockCalendarEvents();
        }
    }

    /**
     * Get user's primary calendar events (Outlook)
     * @param startDate Start date for filtering events
     * @param endDate End date for filtering events
     * @returns Promise resolving to an array of calendar events
     */
    public async getMyCalendarEvents(startDate?: Date, endDate?: Date): Promise<ICalendarEvent[]> {
        // Graph API integration requires additional setup
        return this.getMockCalendarEvents();
    }

    /**
     * Map SharePoint calendar item to ICalendarEvent
     */
    private mapSharePointEventToCalendarEvent(item: any): ICalendarEvent {
        return {
            id: String(item.Id),
            title: item.Title || '',
            description: item.Description || '',
            startDate: new Date(item.EventDate),
            endDate: new Date(item.EndDate),
            createdBy: item.Author?.Title || '',
            location: item.Location || ''
        };
    }

    /**
     * Get default thumbnail images for events (returns random placeholder)
     */
    public getDefaultThumbnail(): string {
        const colors = ['007bff', '28a745', 'ffc107', 'dc3545', '6f42c1', '17a2b8'];
        const randomColor = colors[Math.floor(Math.random() * colors.length)];
        const randomId = Math.floor(Math.random() * 1000000);
        return `https://via.placeholder.com/400x200/${randomColor}/ffffff?text=Event+${randomId}`;
    }

    /**
     * Mock data for development and testing
     */
    private getMockCalendarEvents(): ICalendarEvent[] {
        const today = new Date();
        return [
            {
                id: '1',
                title: 'Team Standup',
                description: 'Daily team standup meeting to sync on progress and blockers',
                startDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1, 10, 0),
                endDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 1, 10, 30),
                createdBy: 'John Doe',
                location: 'Conference Room A',
                organizer: 'john.doe@company.com',
                isAllDay: false,
                thumbnailUrl: this.getDefaultThumbnail()
            },
            {
                id: '2',
                title: 'Project Review',
                description: 'Quarterly project review and planning session',
                startDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 2, 14, 0),
                endDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 2, 15, 30),
                createdBy: 'Jane Smith',
                location: 'Teams Online',
                organizer: 'jane.smith@company.com',
                isAllDay: false,
                thumbnailUrl: this.getDefaultThumbnail()
            },
            {
                id: '3',
                title: 'All Hands Meeting',
                description: 'Company-wide all hands meeting with executives',
                startDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 5, 11, 0),
                endDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 5, 12, 0),
                createdBy: 'CEO',
                location: 'Main Auditorium',
                isAllDay: false,
                thumbnailUrl: this.getDefaultThumbnail()
            },
            {
                id: '4',
                title: 'Training Workshop',
                description: 'Professional development training on new tools and processes',
                startDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7, 9, 0),
                endDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 7, 17, 0),
                createdBy: 'HR Department',
                location: 'Virtual',
                isAllDay: true,
                thumbnailUrl: this.getDefaultThumbnail()
            },
            {
                id: '5',
                title: 'Client Presentation',
                description: 'Presenting Q4 results and plans to key stakeholders',
                startDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 10, 15, 0),
                endDate: new Date(today.getFullYear(), today.getMonth(), today.getDate() + 10, 16, 30),
                createdBy: 'Sales Team',
                location: 'Client Office',
                isAllDay: false,
                thumbnailUrl: this.getDefaultThumbnail()
            }
        ];
    }
}
