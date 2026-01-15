import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IQuickLinksProps {
    context: WebPartContext;
    listTitle: string;
    title?: string;
    department?: string;
}
