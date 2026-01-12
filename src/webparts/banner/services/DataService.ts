export interface IFeaturedItem {
  id: string;
  title: string;
  excerpt: string;
  views: number;
  comments: number;
  category?: string;
}

export default class DataService {
  private context: any;
  private useMock: boolean = true;

  constructor(context?: any) {
    this.context = context;
  }

  public async getCurrentUserDisplayName(): Promise<string> {
    if (this.useMock || !this.context) {

      const displayName: string = this.context.pageContext.user.displayName;
      return Promise.resolve(displayName);
    }

    // Example using SPHttpClient to get current user
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser?$select=Title`;
    const response = await this.context.spHttpClient.get(requestUrl, this.context.spHttpClient.configurations.v1);
    const json = await response.json();
    return json.Title || 'User';
  }

  public async getFeaturedItems(): Promise<IFeaturedItem[]> {
    if (this.useMock || !this.context) {
      return Promise.resolve([
        { id: '1', title: 'Employee Success Stories: Share Your Journey', excerpt: 'Tell us how you grew at the company.', views: 6, comments: 0, category: 'Announcements' }
      ]);
    }

    // Example SharePoint list call (adjust list name and fields)
    const listTitle = 'Featured';
    const requestUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listTitle}')/items?$select=Id,Title,Body,Views,Comments`;
    const response = await this.context.spHttpClient.get(requestUrl, this.context.spHttpClient.configurations.v1);
    const json = await response.json();
    return (json.value || []).map((i: any) => ({
      id: String(i.Id),
      title: i.Title,
      excerpt: i.Body ? (i.Body.length > 120 ? i.Body.substring(0, 117) + '...' : i.Body) : '',
      views: i.Views || 0,
      comments: i.Comments || 0
    }));
  }
}