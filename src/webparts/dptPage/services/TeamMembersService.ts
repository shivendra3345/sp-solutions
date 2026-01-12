import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITeamMember {
    id: string;
    displayName: string;
    jobTitle?: string;
    email?: string;
    phone?: string;
    location?: string;
    photoUrl?: string;
}

export default class TeamMembersService {
    private context: WebPartContext;
    private useMock: boolean = false;

    constructor(context: WebPartContext) {
        this.context = context;
        this.useMock = !context || !context.spHttpClient;
    }

    /**
     * Fetch team members from a SharePoint list
     * @param listTitle The title of the SharePoint list containing team members
     * @returns Promise resolving to an array of team members
     */
    public async getTeamMembers(listTitle: string): Promise<ITeamMember[]> {
        if (this.useMock || !listTitle) {
            return this.getMockTeamMembers();
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            // Query to get team member items
            const requestUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Id,Title,JobTitle,Email,Phone,Location,PhotoUrl&$orderby=Title asc&$top=100`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Failed to fetch team members: ${response.statusText}`);
            }

            const data = await response.json();
            return (data.value || []).map((item: any) => this.mapItemToTeamMember(item));
        } catch (error) {
            console.error('Error fetching team members:', error);
            return this.getMockTeamMembers();
        }
    }

    /**
     * Fetch a single team member by ID
     * @param listTitle The title of the SharePoint list
     * @param itemId The ID of the team member to fetch
     * @returns Promise resolving to a single team member or null
     */
    public async getTeamMemberById(listTitle: string, itemId: number): Promise<ITeamMember | null> {
        if (this.useMock || !listTitle) {
            const mockMembers = this.getMockTeamMembers();
            return mockMembers.find(m => m.id === String(itemId)) || null;
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const requestUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})?$select=Id,Title,JobTitle,Email,Phone,Location,PhotoUrl`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                return null;
            }

            const data = await response.json();
            return this.mapItemToTeamMember(data);
        } catch (error) {
            console.error('Error fetching team member by ID:', error);
            return null;
        }
    }

    /**
     * Helper method to map SharePoint list item to ITeamMember interface
     */
    private mapItemToTeamMember(item: any): ITeamMember {
        return {
            id: String(item.Id),
            displayName: item.Title || '',
            jobTitle: item.JobTitle || '',
            email: item.Email || '',
            phone: item.Phone || '',
            location: item.Location || '',
            photoUrl: item.PhotoUrl || ''
        };
    }

    /**
     * Get default avatar for team member (initials-based placeholder)
     */
    public getDefaultAvatar(displayName: string): string {
        const initials = displayName
            .split(' ')
            .map(n => n[0])
            .join('')
            .toUpperCase()
            .slice(0, 2);

        const colors = ['007bff', '28a745', 'ffc107', 'dc3545', '6f42c1', '17a2b8', 'e83e8c', 'fd7e14'];
        const colorIndex = (displayName.length + initials.charCodeAt(0)) % colors.length;
        const color = colors[colorIndex];

        return `https://via.placeholder.com/200/${color}/ffffff?text=${initials}`;
    }

    /**
     * Mock data for development and testing
     */
    private getMockTeamMembers(): ITeamMember[] {
        return [
            {
                id: '1',
                displayName: 'John Smith',
                jobTitle: 'Senior Developer',
                email: 'john.smith@company.com',
                phone: '+1 (555) 123-4567',
                location: 'New York, NY',
                photoUrl: 'https://via.placeholder.com/200/007bff/ffffff?text=JS'
            },
            {
                id: '2',
                displayName: 'Sarah Johnson',
                jobTitle: 'Product Manager',
                email: 'sarah.johnson@company.com',
                phone: '+1 (555) 234-5678',
                location: 'San Francisco, CA',
                photoUrl: 'https://via.placeholder.com/200/28a745/ffffff?text=SJ'
            },
            {
                id: '3',
                displayName: 'Michael Chen',
                jobTitle: 'UX Designer',
                email: 'michael.chen@company.com',
                phone: '+1 (555) 345-6789',
                location: 'Seattle, WA',
                photoUrl: 'https://via.placeholder.com/200/ffc107/ffffff?text=MC'
            },
            {
                id: '4',
                displayName: 'Emily Davis',
                jobTitle: 'QA Engineer',
                email: 'emily.davis@company.com',
                phone: '+1 (555) 456-7890',
                location: 'Austin, TX',
                photoUrl: 'https://via.placeholder.com/200/dc3545/ffffff?text=ED'
            },
            {
                id: '5',
                displayName: 'James Wilson',
                jobTitle: 'DevOps Engineer',
                email: 'james.wilson@company.com',
                phone: '+1 (555) 567-8901',
                location: 'Denver, CO',
                photoUrl: 'https://via.placeholder.com/200/6f42c1/ffffff?text=JW'
            },
            {
                id: '6',
                displayName: 'Lisa Anderson',
                jobTitle: 'Business Analyst',
                email: 'lisa.anderson@company.com',
                phone: '+1 (555) 678-9012',
                location: 'Boston, MA',
                photoUrl: 'https://via.placeholder.com/200/17a2b8/ffffff?text=LA'
            }
        ];
    }
}
