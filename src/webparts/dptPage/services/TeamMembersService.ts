import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ITeamMember {
    id: string;
    displayName: string;
    jobTitle?: string;
    email?: string;
    phone?: string;
    location?: string;
    department?: string;
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
    public async getTeamMembers(listTitle: string, department?: string): Promise<ITeamMember[]> {
        if (this.useMock || !listTitle) {
            return this.getMockTeamMembers();
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;

            // Query items where IsActive is true and expand the Person field 'Name'
            // Assumes the list has: 'Name' (Person or Group) and 'IsActive' (Yes/No)
            // Note: querying Name/LoginName directly is not supported; select Id/Title/EMail and expand Name
            // Also select Department if present; we'll optionally filter by department if provided
            let selectFields = 'Id,Name/Id,Name/Title,Name/EMail,IsActive';
            if (department) {
                selectFields += ',Department';
            } else {
                // still attempt to include Department in the response when available
                selectFields += ',Department';
            }
            let requestUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=${selectFields}&$expand=Name&$filter=IsActive eq 1`;
            if (department) {
                // filter items whose Department equals the provided department (case-sensitive match)
                requestUrl += ` and Department eq '${encodeURIComponent(department)}'`;
            }
            requestUrl += '&$orderby=Name/Title asc&$top=500';

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Failed to fetch team members: ${response.statusText}`);
            }

            const data = await response.json();
            const items: any[] = data.value || [];

            // Cache user profile lookups to avoid duplicate PeopleManager calls
            const profileCache: Map<string, any> = new Map();

            const members: ITeamMember[] = [];

            for (const item of items) {
                // If the person field is empty, skip
                const person = item.Name;
                if (!person || (person.Id === undefined || person.Id === null)) {
                    continue;
                }

                // Resolve the user's login name via getUserById since Name/LoginName is not selectable
                const userId = Number(person.Id);
                let loginName: string | null = await this.getUserLoginName(userId);

                // Fallback to email from the person field if LoginName couldn't be resolved
                if (!loginName && person.EMail) {
                    loginName = person.EMail;
                }

                // Get profile info (cached) if we have a loginName
                let profileProps = null;
                if (loginName) {
                    profileProps = profileCache.get(loginName);
                    if (!profileProps) {
                        profileProps = await this.getUserProfileProperties(loginName);
                        profileCache.set(loginName, profileProps);
                    }
                }

                const displayName = person.Title || (profileProps && profileProps.DisplayName) || '';
                const email = person.EMail || (profileProps && (profileProps.WorkEmail || profileProps.Email)) || '';
                const jobTitle = (profileProps && (profileProps.SPSJobTitle || profileProps.JobTitle || profileProps.Title)) || '';
                const location = (profileProps && (profileProps.SPSLocation || profileProps.Office || profileProps.Location)) || '';

                // Use PeopleManager picture URL if available; otherwise userphoto.aspx with accountname fallback
                const photoUrl = (profileProps && (profileProps.PictureURL || profileProps.UserProfileProfilePicture || profileProps.Picture)) ||
                    `${webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(loginName || person.EMail || '')}`;

                const phone = (profileProps && (profileProps.Phone || profileProps.CellPhone || profileProps.WorkPhone || profileProps.Phone)) || '';

                members.push({
                    id: String(item.Id),
                    displayName,
                    jobTitle,
                    email,
                    phone,
                    location,
                    department: item.Department || undefined,
                    photoUrl
                });
            }

            return members;
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
            const requestUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})?$select=Id,Name/Id,Name/Title,Name/EMail,IsActive,Department&$expand=Name`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                return null;
            }

            const data = await response.json();
            const item = data;
            const person = item.Name;
            if (!person || (person.Id === undefined || person.Id === null)) {
                return null;
            }

            const userId = Number(person.Id);
            let loginName: string | null = await this.getUserLoginName(userId);
            if (!loginName && person.EMail) {
                loginName = person.EMail;
            }

            const profile = loginName ? await this.getUserProfileProperties(loginName) : null;

            const displayName = person.Title || (profile && profile.DisplayName) || '';
            const email = person.EMail || (profile && (profile.WorkEmail || profile.Email)) || '';
            const jobTitle = (profile && (profile.SPSJobTitle || profile.JobTitle || profile.Title)) || '';
            const location = (profile && (profile.SPSLocation || profile.Office || profile.Location)) || '';
            const photoUrl = (profile && (profile.PictureURL || profile.UserProfileProfilePicture || profile.Picture)) ||
                `${webUrl}/_layouts/15/userphoto.aspx?size=M&accountname=${encodeURIComponent(loginName || person.EMail || '')}`;

            const phone = (profile && (profile.Phone || profile.CellPhone || profile.WorkPhone || profile.Phone)) || '';

            return {
                id: String(item.Id),
                displayName,
                jobTitle,
                email,
                phone,
                location,
                department: item.Department || undefined,
                photoUrl
            };
        } catch (error) {
            console.error('Error fetching team member by ID:', error);
            return null;
        }
    }

    /**
     * Helper method to map SharePoint list item to ITeamMember interface
     */
    // private mapItemToTeamMember(item: any): ITeamMember {
    //     // Keep old mapping as fallback for legacy lists
    //     return {
    //         id: String(item.Id),
    //         displayName: item.Title || (item.Name && item.Name.Title) || '',
    //         jobTitle: item.JobTitle || '',
    //         email: item.Email || (item.Name && item.Name.EMail) || '',
    //         phone: item.Phone || '',
    //         location: item.Location || '',
    //         photoUrl: item.PhotoUrl || ''
    //     };
    // }

    /**
     * Fetch user profile properties from PeopleManager for a given account (loginName).
     * Returns an object with convenient keys (DisplayName, WorkEmail, SPSJobTitle, SPSLocation, PictureURL, Office ...)
     */
    private async getUserProfileProperties(loginName: string): Promise<any> {
        if (!loginName || !this.context || !this.context.spHttpClient) {
            return null;
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            // PeopleManager expects the accountName parameter wrapped in single quotes and URL encoded
            const encoded = encodeURIComponent(loginName);
            const peopleManagerUrl = `${webUrl}/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='${encoded}'`;

            const resp: SPHttpClientResponse = await this.context.spHttpClient.get(
                peopleManagerUrl,
                SPHttpClient.configurations.v1
            );

            if (!resp.ok) {
                return null;
            }

            const profile = await resp.json();

            // Convert UserProfileProperties array to a dictionary for easy lookup
            const props: any = {};
            if (profile && Array.isArray(profile.UserProfileProperties)) {
                for (const p of profile.UserProfileProperties) {
                    if (p && p.Key) {
                        props[p.Key] = p.Value;
                    }
                }
            }

            // Add some convenient aliases
            props.DisplayName = profile && profile.DisplayName;
            props.AccountName = profile && profile.AccountName;
            props.WorkEmail = props['WorkEmail'] || props['SPS-WorkEmail'] || props['WorkEmail'] || props['Email'];
            // Phone aliases - PeopleManager may use different property names across tenants
            props.WorkPhone = props['WorkPhone'] || props['WorkTelephone'] || props['TelephoneNumber'] || props['Phone'] || null;
            props.CellPhone = props['CellPhone'] || props['Mobile'] || props['Cellphone'] || props['MobilePhone'] || null;
            props.Phone = props.WorkPhone || props.CellPhone || props['HomePhone'] || props['TelephoneNumber'] || null;
            props.SPSJobTitle = props['SPS-JobTitle'] || props['SPSJobTitle'] || props['JobTitle'];
            props.SPSLocation = props['SPS-Location'] || props['SPSLocation'] || props['Location'] || props['Office'];
            props.PictureURL = props['PictureURL'] || props['Picture'] || props['UserProfileProfilePicture'];

            return props;
        } catch (e) {
            console.error('Error fetching user profile for', loginName, e);
            return null;
        }
    }

    /**
     * Resolve a user's LoginName by SharePoint user id using /getUserById
     */
    private async getUserLoginName(userId: number): Promise<string | null> {
        if (!userId || !this.context || !this.context.spHttpClient) {
            return null;
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const url = `${webUrl}/_api/web/getUserById(${userId})?$select=Id,LoginName,Email,Title`;
            const resp: SPHttpClientResponse = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
            if (!resp.ok) {
                return null;
            }

            const data = await resp.json();
            return data && data.LoginName ? data.LoginName : null;
        } catch (e) {
            console.error('Error resolving user loginName for id', userId, e);
            return null;
        }
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
