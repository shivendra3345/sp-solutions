import { WebPartContext } from '@microsoft/sp-webpart-base';
import { IBannerImage } from '../../services/BannerImageService';

export type BannerLayout = 'carousel' | 'tile';

export interface IPageBannerProps {
    context: WebPartContext;
    layout: BannerLayout;
    listTitle: string;
    title?: string;
    isDarkTheme?: boolean;
}

export interface IPageBannerState {
    images: IBannerImage[];
    currentIndex: number;
    loading: boolean;
    error?: string;
}
