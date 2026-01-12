import * as React from 'react';
import styles from './PageBanner.module.scss';
import { IPageBannerProps, IPageBannerState } from '../PageBanner/IPageBannerProps';
import BannerImageService, { IBannerImage } from '../../services/BannerImageService';

export default class PageBanner extends React.Component<IPageBannerProps, IPageBannerState> {
    private bannerService: BannerImageService;

    constructor(props: IPageBannerProps) {
        super(props);
        this.bannerService = new BannerImageService(props.context);
        this.state = {
            images: [],
            currentIndex: 0,
            loading: true
        };
    }

    public componentDidMount(): void {
        this.loadBannerImages();
    }

    public componentDidUpdate(prevProps: IPageBannerProps): void {
        if (prevProps.listTitle !== this.props.listTitle || prevProps.layout !== this.props.layout) {
            this.loadBannerImages();
        }
    }

    private loadBannerImages(): void {
        this.setState({ loading: true });
        this.bannerService.getBannerImages(this.props.listTitle)
            .then((images: IBannerImage[]) => {
                this.setState({
                    images,
                    currentIndex: 0,
                    loading: false
                });
            })
            .catch((error: Error) => {
                console.error('Error loading banner images:', error);
                this.setState({
                    error: error.message,
                    loading: false
                });
            });
    }

    private handleNextSlide = (): void => {
        const { images, currentIndex } = this.state;
        if (images.length === 0) return;
        this.setState({
            currentIndex: (currentIndex + 1) % images.length
        });
    };

    private handlePrevSlide = (): void => {
        const { images, currentIndex } = this.state;
        if (images.length === 0) return;
        this.setState({
            currentIndex: (currentIndex - 1 + images.length) % images.length
        });
    };

    private handleDotClick = (index: number): void => {
        this.setState({ currentIndex: index });
    };

    private renderCarouselLayout(): React.ReactElement {
        const { images, currentIndex, loading } = this.state;

        if (loading) {
            return <div className={styles.loading}>Loading banner...</div>;
        }

        if (images.length === 0) {
            return <div className={styles.noImages}>No images available</div>;
        }

        const currentImage = images[currentIndex];

        return (
            <div className={styles.carousel}>
                <div className={styles.carouselContainer}>
                    <img
                        src={currentImage.imageUrl}
                        alt={currentImage.title}
                        className={styles.carouselImage}
                    />
                    <div className={styles.carouselOverlay}>
                        <h2 className={styles.carouselTitle}>{currentImage.title}</h2>
                        {currentImage.description && (
                            <p className={styles.carouselDescription}>{currentImage.description}</p>
                        )}
                        {currentImage.link && (
                            <a href={currentImage.link} className={styles.carouselLink}>
                                Learn More →
                            </a>
                        )}
                    </div>

                    {images.length > 1 && (
                        <>
                            <button
                                className={`${styles.carouselButton} ${styles.prevButton}`}
                                onClick={this.handlePrevSlide}
                                aria-label="Previous slide"
                            >
                                &#10094;
                            </button>
                            <button
                                className={`${styles.carouselButton} ${styles.nextButton}`}
                                onClick={this.handleNextSlide}
                                aria-label="Next slide"
                            >
                                &#10095;
                            </button>
                        </>
                    )}
                </div>

                {images.length > 1 && (
                    <div className={styles.dotsContainer}>
                        {images.map((_: IBannerImage, index: number) => (
                            <button
                                key={index}
                                className={`${styles.dot} ${index === currentIndex ? styles.activeDot : ''}`}
                                onClick={() => this.handleDotClick(index)}
                                aria-label={`Go to slide ${index + 1}`}
                            />
                        ))}
                    </div>
                )}
            </div>
        );
    }

    private renderTileLayout(): React.ReactElement {
        const { images, loading } = this.state;

        if (loading) {
            return <div className={styles.loading}>Loading banner...</div>;
        }

        if (images.length === 0) {
            return <div className={styles.noImages}>No images available</div>;
        }

        // For tile layout: 1 large image on left, 2 smaller images on right
        const mainImage = images[0];
        const tileImages = images.slice(1, 3);

        return (
            <div className={styles.tileContainer}>
                {/* Main large tile */}
                <div className={styles.mainTile}>
                    <img
                        src={mainImage.imageUrl}
                        alt={mainImage.title}
                        className={styles.mainTileImage}
                    />
                    <div className={styles.tileOverlay}>
                        <h2 className={styles.tileTitle}>{mainImage.title}</h2>
                        {mainImage.description && (
                            <p className={styles.tileDescription}>{mainImage.description}</p>
                        )}
                        {mainImage.link && (
                            <a href={mainImage.link} className={styles.tileLink}>
                                Learn More →
                            </a>
                        )}
                    </div>
                </div>

                {/* Smaller tiles on the right */}
                <div className={styles.smallTilesColumn}>
                    {tileImages.map((image: IBannerImage, index: number) => (
                        <div key={index} className={styles.smallTile}>
                            <img
                                src={image.imageUrl}
                                alt={image.title}
                                className={styles.smallTileImage}
                            />
                            <div className={styles.smallTileOverlay}>
                                <h3 className={styles.smallTileTitle}>{image.title}</h3>
                                {image.description && (
                                    <p className={styles.smallTileDescription}>{image.description}</p>
                                )}
                            </div>
                        </div>
                    ))}
                    {tileImages.length < 2 && (
                        // Add empty placeholder if less than 2 images
                        Array.from({ length: 2 - tileImages.length }).map((_: any, index: number) => (
                            <div key={`empty-${index}`} className={styles.smallTile}>
                                <div className={styles.emptyTile}>No image</div>
                            </div>
                        ))
                    )}
                </div>
            </div>
        );
    }

    public render(): React.ReactElement<IPageBannerProps> {
        const { layout } = this.props;
        const { error } = this.state;

        if (error) {
            return (
                <div className={styles.error}>
                    <p>Error loading banner: {error}</p>
                </div>
            );
        }

        return (
            <div className={`${styles.pageBanner} ${this.props.isDarkTheme ? styles.darkTheme : ''}`}>
                {layout === 'carousel' ? this.renderCarouselLayout() : this.renderTileLayout()}
            </div>
        );
    }
}
