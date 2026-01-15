import * as React from 'react';
import styles from './QuickLinks.module.scss';
import { IQuickLinksProps } from './IQuickLinksProps';
import QuickLinksService, { IQuickLink } from '../../services/QuickLinksService';

interface IQuickLinksState {
    quickLinks: IQuickLink[];
    currentIndex: number;
    loading: boolean;
    error: undefined | string;
    itemsPerSlide: number;
}

export default class QuickLinks extends React.Component<IQuickLinksProps, IQuickLinksState> {
    private quickLinksService: QuickLinksService;
    private resizeTimer: number | null = null;

    constructor(props: IQuickLinksProps) {
        super(props);
        this.quickLinksService = new QuickLinksService(props.context);
        this.state = {
            quickLinks: [],
            currentIndex: 0,
            loading: true,
            error: undefined,
            itemsPerSlide: this.getItemsPerSlide()
        };
    }

    public componentDidMount(): void {
        this.loadQuickLinks();
        window.addEventListener('resize', this.handleWindowResize);
    }

    public componentWillUnmount(): void {
        window.removeEventListener('resize', this.handleWindowResize);
        if (this.resizeTimer) {
            clearTimeout(this.resizeTimer);
        }
    }

    public componentDidUpdate(prevProps: IQuickLinksProps): void {
        if (prevProps.listTitle !== this.props.listTitle) {
            this.loadQuickLinks();
        }
    }

    private loadQuickLinks(): void {
        this.setState({ loading: true, error: undefined });

        this.quickLinksService.getQuickLinks(this.props.listTitle, this.props.department)
            .then((quickLinks: IQuickLink[]) => {
                this.setState({
                    quickLinks,
                    loading: false,
                    currentIndex: 0
                });
            })
            .catch((error: Error) => {
                console.error('Error loading quick links:', error);
                this.setState({
                    loading: false,
                    error: error.message || 'Failed to load quick links'
                });
            });
    }

    private getItemsPerSlide(): number {
        const width = typeof window !== 'undefined' ? window.innerWidth : 1200;
        if (width < 768) return 2;
        if (width < 1024) return 4;
        return 6;
    }

    private handleWindowResize = (): void => {
        if (this.resizeTimer) {
            clearTimeout(this.resizeTimer);
        }
        this.resizeTimer = window.setTimeout(() => {
            const newItemsPerSlide = this.getItemsPerSlide();
            if (newItemsPerSlide !== this.state.itemsPerSlide) {
                this.setState({ itemsPerSlide: newItemsPerSlide, currentIndex: 0 });
            }
        }, 250);
    };

    private handlePreviousClick = (): void => {
        const { currentIndex, itemsPerSlide } = this.state;
        const newIndex = Math.max(0, currentIndex - itemsPerSlide);
        this.setState({ currentIndex: newIndex });
    };

    private handleNextClick = (): void => {
        const { currentIndex, quickLinks, itemsPerSlide } = this.state;
        const maxIndex = Math.max(0, quickLinks.length - itemsPerSlide);
        const newIndex = Math.min(maxIndex, currentIndex + itemsPerSlide);
        this.setState({ currentIndex: newIndex });
    };

    private handleQuickLinkClick = (link: string): void => {
        if (link) {
            window.open(link, '_blank');
        }
    };

    private canGoNext = (): boolean => {
        const { currentIndex, quickLinks, itemsPerSlide } = this.state;
        const maxIndex = Math.max(0, quickLinks.length - itemsPerSlide);
        return currentIndex < maxIndex;
    };

    private canGoPrevious = (): boolean => {
        return this.state.currentIndex > 0;
    };

    public render(): React.ReactElement<IQuickLinksProps> {
        const { quickLinks, currentIndex, loading, error, itemsPerSlide } = this.state;

        if (loading) {
            return <div className={styles.quickLinksContainer}>Loading quick links...</div>;
        }

        if (error) {
            return <div className={styles.quickLinksContainer}>Error: {error}</div>;
        }

        if (quickLinks.length === 0) {
            return <div className={styles.quickLinksContainer}>No quick links available</div>;
        }

        const visibleLinks = quickLinks.slice(currentIndex, currentIndex + itemsPerSlide);

        return (
            <div className={styles.quickLinksContainer}>
                {this.props.title && (
                    <div className={styles.quickLinksHeader}>
                        <h2 className={styles.quickLinksTitle}>{this.props.title}</h2>
                    </div>
                )}

                <div className={styles.carouselWrapper}>
                    <button
                        className={`${styles.navButton} ${styles.prevButton} ${!this.canGoPrevious() ? styles.disabled : ''}`}
                        onClick={this.handlePreviousClick}
                        disabled={!this.canGoPrevious()}
                        aria-label="Previous"
                    >
                        &#10094;
                    </button>

                    <div className={styles.carouselTrack}>
                        {visibleLinks.map((link: IQuickLink, index: number) => (
                            <div
                                key={link.id}
                                className={styles.quickLinkCard}
                                onClick={() => this.handleQuickLinkClick(link.link)}
                                role="button"
                                tabIndex={0}
                                onKeyPress={(e: React.KeyboardEvent) => {
                                    if (e.key === 'Enter' || e.key === ' ') {
                                        this.handleQuickLinkClick(link.link);
                                    }
                                }}
                            >
                                <div
                                    className={styles.cardImage}
                                    style={{ backgroundImage: `url('${link.imageUrl}')` }}
                                />
                                <div className={styles.cardOverlay}>
                                    <span className={styles.cardTitle}>{link.title}</span>
                                </div>
                            </div>
                        ))}
                    </div>

                    <button
                        className={`${styles.navButton} ${styles.nextButton} ${!this.canGoNext() ? styles.disabled : ''}`}
                        onClick={this.handleNextClick}
                        disabled={!this.canGoNext()}
                        aria-label="Next"
                    >
                        &#10095;
                    </button>
                </div>

                {quickLinks.length > itemsPerSlide && (
                    <div className={styles.indicators}>
                        {Array.from({ length: Math.ceil(quickLinks.length / itemsPerSlide) }, (_, i: number) => (
                            <div
                                key={i}
                                className={`${styles.indicator} ${i === Math.floor(currentIndex / itemsPerSlide) ? styles.active : ''}`}
                            />
                        ))}
                    </div>
                )}
            </div>
        );
    }
}
