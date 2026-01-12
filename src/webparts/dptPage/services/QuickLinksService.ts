import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IQuickLink {
    id: string;
    title: string;
    imageUrl: string;
    link: string;
}

export default class QuickLinksService {
    private context: WebPartContext;
    private useMock: boolean = false;

    constructor(context: WebPartContext) {
        this.context = context;
        this.useMock = !context || !context.spHttpClient;
    }

    /**
     * Fetch quick links from a SharePoint list
     * @param listTitle The title or name of the SharePoint list containing quick links
     * @returns Promise resolving to an array of quick links
     */
    public async getQuickLinks(listTitle: string): Promise<IQuickLink[]> {
        if (this.useMock) {
            return this.getMockQuickLinks();
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            // Query to get items with image URL, title, and link fields
            const requestUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Id,Title,ImageUrl,Link&$orderby=Created desc`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Failed to fetch quick links: ${response.statusText}`);
            }

            const data = await response.json();
            return (data.value || []).map((item: any) => this.mapItemToQuickLink(item));
        } catch (error) {
            console.error('Error fetching quick links:', error);
            return this.getMockQuickLinks();
        }
    }

    private mapItemToQuickLink(item: any): IQuickLink {
        return {
            id: String(item.Id),
            title: item.Title || '',
            imageUrl: item.ImageUrl || '',
            link: item.Link || ''
        };
    }

    private getMockQuickLinks(): IQuickLink[] {
        return [
            {
                id: '1',
                title: 'Documentation',
                imageUrl: 'https://via.placeholder.com/150/0078d4/ffffff?text=Docs',
                link: 'https://example.com/docs'
            },
            {
                id: '2',
                title: 'Support',
                imageUrl: 'https://via.placeholder.com/150/107c10/ffffff?text=Support',
                link: 'https://example.com/support'
            },
            {
                id: '3',
                title: 'Training',
                imageUrl: 'https://via.placeholder.com/150/7030a0/ffffff?text=Training',
                link: 'https://example.com/training'
            },
            {
                id: '4',
                title: 'Updates',
                imageUrl: 'https://via.placeholder.com/150/ff8c00/ffffff?text=Updates',
                link: 'https://example.com/updates'
            },
            {
                id: '5',
                title: 'Community',
                imageUrl: 'https://via.placeholder.com/150/038387/ffffff?text=Community',
                link: 'https://example.com/community'
            },
            {
                id: '6',
                title: 'Resources',
                imageUrl: 'https://via.placeholder.com/150/7030a0/ffffff?text=Resources',
                link: 'https://example.com/resources'
            }
        ];
    }
}
