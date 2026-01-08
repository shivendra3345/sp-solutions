import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IBannerImage {
    id: string;
    title: string;
    imageUrl: string;
    description?: string;
    link?: string;
}

export default class BannerImageService {
    private context: WebPartContext;
    private useMock: boolean = false;

    constructor(context: WebPartContext) {
        this.context = context;
        this.useMock = !context || !context.spHttpClient;
    }

    /**
     * Fetch banner images from a SharePoint list
     * @param listTitle The title or name of the SharePoint list containing banner images
     * @returns Promise resolving to an array of banner images
     */
    public async getBannerImages(listTitle: string): Promise<IBannerImage[]> {
        if (this.useMock) {
            return this.getMockBannerImages();
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            // Query to get items with image URL, title, description, and link fields
            const requestUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items?$select=Id,Title,ImageUrl,Description,Link&$orderby=Created desc`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                throw new Error(`Failed to fetch banner images: ${response.statusText}`);
            }

            const data = await response.json();
            return (data.value || []).map((item: any) => this.mapItemToBannerImage(item));
        } catch (error) {
            console.error('Error fetching banner images:', error);
            return this.getMockBannerImages();
        }
    }

    /**
     * Fetch a single banner image by ID
     * @param listTitle The title of the SharePoint list
     * @param itemId The ID of the item to fetch
     * @returns Promise resolving to a single banner image or null
     */
    public async getBannerImageById(listTitle: string, itemId: number): Promise<IBannerImage | null> {
        if (this.useMock) {
            const mockImages = this.getMockBannerImages();
            return mockImages.find(img => img.id === String(itemId)) || null;
        }

        try {
            const webUrl = this.context.pageContext.web.absoluteUrl;
            const requestUrl = `${webUrl}/_api/web/lists/getbytitle('${encodeURIComponent(listTitle)}')/items(${itemId})?$select=Id,Title,ImageUrl,Description,Link`;

            const response: SPHttpClientResponse = await this.context.spHttpClient.get(
                requestUrl,
                SPHttpClient.configurations.v1
            );

            if (!response.ok) {
                return null;
            }

            const data = await response.json();
            return this.mapItemToBannerImage(data);
        } catch (error) {
            console.error('Error fetching banner image by ID:', error);
            return null;
        }
    }

    /**
     * Helper method to map SharePoint list item to IBannerImage interface
     */
    private mapItemToBannerImage(item: any): IBannerImage {
        return {
            id: String(item.Id),
            title: item.Title || '',
            imageUrl: item.ImageUrl || '',
            description: item.Description || '',
            link: item.Link || ''
        };
    }

    /**
     * Mock data for development and testing
     */
    private getMockBannerImages(): IBannerImage[] {
        return [
            {
                id: '1',
                title: 'Welcome to SharePoint',
                imageUrl: 'https://via.placeholder.com/1200x300/007bff/ffffff?text=SharePoint+Banner',
                description: 'Your collaborative platform',
                link: '/'
            },
            {
                id: '2',
                title: 'Team Collaboration',
                imageUrl: 'https://via.placeholder.com/400x300/28a745/ffffff?text=Collaboration',
                description: 'Work together effectively',
                link: '/sites/collaboration'
            },
            {
                id: '3',
                title: 'Knowledge Sharing',
                imageUrl: 'https://via.placeholder.com/400x300/ffc107/ffffff?text=Knowledge',
                description: 'Share and learn from others',
                link: '/sites/knowledge'
            }
        ];
    }
}
