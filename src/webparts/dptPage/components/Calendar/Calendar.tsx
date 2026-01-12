import * as React from 'react';
import styles from './Calendar.module.scss';
import { ICalendarProps, ICalendarState } from '../Calendar/ICalendarProps';
import CalendarEventService, { ICalendarEvent } from '../../services/CalendarEventService';

export default class Calendar extends React.Component<ICalendarProps, ICalendarState> {
    private calendarService: CalendarEventService;
    private readonly DAYS_OF_WEEK = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];

    constructor(props: ICalendarProps) {
        super(props);
        this.calendarService = new CalendarEventService(props.context);
        this.state = {
            events: [],
            filteredEvents: [],
            currentIndex: 0,
            loading: true,
            selectedDate: new Date(),
            layout: this.props.layout || 'grid',
            error: undefined
        };
    }

    public componentDidMount(): void {
        this.loadCalendarEvents();
    }

    public componentDidUpdate(prevProps: ICalendarProps): void {
        if (
            prevProps.calendarSource !== this.props.calendarSource ||
            prevProps.groupId !== this.props.groupId ||
            prevProps.spListTitle !== this.props.spListTitle
        ) {
            this.loadCalendarEvents();
        }
    }

    private loadCalendarEvents(): void {
        this.setState({ loading: true });

        Promise.all([
            this.props.calendarSource === 'outlook' || this.props.calendarSource === 'both'
                ? this.calendarService.getGroupCalendarEvents(this.props.groupId || '')
                : Promise.resolve([]),
            this.props.calendarSource === 'sharepoint' || this.props.calendarSource === 'both'
                ? this.calendarService.getSharePointCalendarEvents(
                    this.props.spSiteUrl || this.props.context.pageContext.web.absoluteUrl,
                    this.props.spListTitle || 'Calendar'
                )
                : Promise.resolve([])
        ])
            .then(([outlookEvents, spEvents]) => {
                const allEvents = [...outlookEvents, ...spEvents];
                // Sort by start date
                allEvents.sort((a, b) => a.startDate.getTime() - b.startDate.getTime());

                this.setState({
                    events: allEvents,
                    filteredEvents: allEvents,
                    loading: false,
                    currentIndex: 0
                });
            })
            .catch((error: Error) => {
                console.error('Error loading calendar events:', error);
                this.setState({
                    error: error.message,
                    loading: false
                });
            });
    }

    private handlePrevMonth = (): void => {
        const { selectedDate } = this.state;
        const newDate = new Date(selectedDate!);
        newDate.setMonth(newDate.getMonth() - 1);
        this.setState({ selectedDate: newDate });
    };

    private handleNextMonth = (): void => {
        const { selectedDate } = this.state;
        const newDate = new Date(selectedDate!);
        newDate.setMonth(newDate.getMonth() + 1);
        this.setState({ selectedDate: newDate });
    };

    private handlePrevCarousel = (): void => {
        const { filteredEvents, currentIndex } = this.state;
        if (filteredEvents.length === 0) return;
        this.setState({
            currentIndex: (currentIndex - 1 + filteredEvents.length) % filteredEvents.length
        });
    };

    private handleNextCarousel = (): void => {
        const { filteredEvents, currentIndex } = this.state;
        if (filteredEvents.length === 0) return;
        this.setState({
            currentIndex: (currentIndex + 1) % filteredEvents.length
        });
    };

    private handleCarouselDotClick = (index: number): void => {
        this.setState({ currentIndex: index });
    };

    private getCalendarDays(): Array<{ date: number; isCurrentMonth: boolean; events: ICalendarEvent[] }> {
        const { selectedDate, events } = this.state;
        const date = selectedDate || new Date();
        const year = date.getFullYear();
        const month = date.getMonth();

        const firstDay = new Date(year, month, 1).getDay();
        const daysInMonth = new Date(year, month + 1, 0).getDate();
        const daysInPrevMonth = new Date(year, month, 0).getDate();

        const days: Array<{ date: number; isCurrentMonth: boolean; events: ICalendarEvent[] }> = [];

        // Previous month days
        for (let i = firstDay - 1; i >= 0; i--) {
            days.push({
                date: daysInPrevMonth - i,
                isCurrentMonth: false,
                events: []
            });
        }

        // Current month days
        for (let i = 1; i <= daysInMonth; i++) {
            const dayEvents = events.filter((event: ICalendarEvent) => {
                const eventDate = new Date(event.startDate);
                return (
                    eventDate.getDate() === i &&
                    eventDate.getMonth() === month &&
                    eventDate.getFullYear() === year
                );
            });

            days.push({
                date: i,
                isCurrentMonth: true,
                events: dayEvents
            });
        }

        // Next month days
        const remainingDays = 42 - days.length; // 6 rows * 7 days
        for (let i = 1; i <= remainingDays; i++) {
            days.push({
                date: i,
                isCurrentMonth: false,
                events: []
            });
        }

        return days;
    }

    private formatEventTime(event: ICalendarEvent): string {
        if (event.isAllDay) {
            return 'All Day';
        }
        const startTime = event.startDate.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit'
        });
        const endTime = event.endDate.toLocaleTimeString('en-US', {
            hour: '2-digit',
            minute: '2-digit'
        });
        return `${startTime} - ${endTime}`;
    }

    private formatDateRange(event: ICalendarEvent): string {
        const startDate = event.startDate.toLocaleDateString('en-US', {
            month: 'short',
            day: 'numeric',
            year: 'numeric'
        });
        const endDate = event.endDate.toLocaleDateString('en-US', {
            month: 'short',
            day: 'numeric',
            year: 'numeric'
        });

        if (startDate === endDate) {
            return startDate;
        }
        return `${startDate} - ${endDate}`;
    }

    private renderCalendarGridLayout(): React.ReactElement {
        const { selectedDate } = this.state;
        const calendarDays = this.getCalendarDays();
        const monthYear = (selectedDate || new Date()).toLocaleString('en-US', { month: 'long', year: 'numeric' });

        return (
            <div>
                <div className={styles.calendarHeader}>
                    <h2 className={styles.calendarTitle}>{this.props.title || 'Calendar'}</h2>
                    <div className={styles.headerControls}>
                        <button className={styles.navButton} onClick={this.handlePrevMonth}>
                            ← Previous
                        </button>
                        <div className={styles.monthYearDisplay}>{monthYear}</div>
                        <button className={styles.navButton} onClick={this.handleNextMonth}>
                            Next →
                        </button>
                    </div>
                </div>

                <div className={styles.calendarGrid}>
                    <div className={styles.monthCell}>
                        {/* Day headers */}
                        {this.DAYS_OF_WEEK.map(day => (
                            <div key={day} className={styles.dayHeader}>
                                {day}
                            </div>
                        ))}

                        {/* Calendar days */}
                        {calendarDays.map((day, index) => {
                            const isToday =
                                day.isCurrentMonth &&
                                day.date === new Date().getDate() &&
                                selectedDate &&
                                selectedDate.getMonth() === new Date().getMonth() &&
                                selectedDate.getFullYear() === new Date().getFullYear();

                            return (
                                <div
                                    key={index}
                                    className={`${styles.dayCell} ${!day.isCurrentMonth ? styles.otherMonth : ''} ${isToday ? styles.today : ''}`}
                                >
                                    <div className={styles.dayCellDate}>{day.date}</div>
                                    <div className={styles.dayCellEvents}>
                                        {day.events.slice(0, 3).map((event: ICalendarEvent, i: number) => (
                                            <div key={i} className={styles.dayCellEvent} title={event.title}>
                                                {event.title}
                                            </div>
                                        ))}
                                        {day.events.length > 3 && (
                                            <div className={styles.moreEvents}>+{day.events.length - 3} more</div>
                                        )}
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                </div>
            </div>
        );
    }

    private renderCarouselLayout(): React.ReactElement {
        const { filteredEvents, currentIndex, loading } = this.state;

        if (loading) {
            return <div className={styles.loading}><div className={styles.spinner}></div></div>;
        }

        if (filteredEvents.length === 0) {
            return <div className={styles.noEvents}><p>No events available</p></div>;
        }

        // Show 3 events at a time in carousel (current + next 2)
        const visibleEvents = [];
        for (let i = 0; i < 3 && i < filteredEvents.length; i++) {
            const eventIndex = (currentIndex + i) % filteredEvents.length;
            visibleEvents.push({ event: filteredEvents[eventIndex], index: eventIndex });
        }

        return (
            <div>
                <div className={styles.calendarHeader}>
                    <h2 className={styles.calendarTitle}>{this.props.title || 'Upcoming Events'}</h2>
                </div>

                <div className={styles.carouselContainer}>
                    <div className={styles.carouselContent}>
                        {visibleEvents.map(({ event, index }: { event: ICalendarEvent; index: number }) => {
                            const thumbnailUrl = event.thumbnailUrl || this.calendarService.getDefaultThumbnail();
                            return (
                                <div
                                    key={index}
                                    className={`${styles.eventCardWrapper} ${index === currentIndex ? styles.active : ''}`}
                                >
                                    <div className={styles.cardImage}>
                                        <img src={thumbnailUrl} alt={event.title} />
                                    </div>

                                    <div className={styles.cardContent}>
                                        <h3 className={styles.cardTitle}>{event.title}</h3>

                                        <div className={styles.cardDateTime}>
                                            {this.formatDateRange(event)} • {this.formatEventTime(event)}
                                        </div>

                                        {event.description && (
                                            <div className={styles.cardDescription}>
                                                {event.description}
                                            </div>
                                        )}

                                        {(event.location || event.createdBy) && (
                                            <div className={styles.cardMeta}>
                                                {event.location && <p><strong>Location:</strong> {event.location}</p>}
                                                {event.createdBy && <p><strong>Organizer:</strong> {event.createdBy}</p>}
                                            </div>
                                        )}
                                    </div>
                                </div>
                            );
                        })}
                    </div>

                    <div className={styles.carouselControls}>
                        <button
                            className={styles.carouselButton}
                            onClick={this.handlePrevCarousel}
                            disabled={filteredEvents.length <= 1}
                            aria-label="Previous events"
                        >
                            ←
                        </button>

                        <div className={styles.carouselIndicators}>
                            {filteredEvents.map((_: ICalendarEvent, index: number) => (
                                <button
                                    key={index}
                                    className={`${styles.indicator} ${index === currentIndex ? styles.active : ''}`}
                                    onClick={() => this.handleCarouselDotClick(index)}
                                    aria-label={`Go to event ${index + 1}`}
                                />
                            ))}
                        </div>

                        <button
                            className={styles.carouselButton}
                            onClick={this.handleNextCarousel}
                            disabled={filteredEvents.length <= 1}
                            aria-label="Next events"
                        >
                            →
                        </button>

                        <span className={styles.pageCounter}>
                            {currentIndex + 1} - {Math.min(currentIndex + 3, filteredEvents.length)} of {filteredEvents.length}
                        </span>
                    </div>
                </div>
            </div>
        );
    }

    public render(): React.ReactElement<ICalendarProps> {
        const { layout, error, loading } = this.state;

        if (error) {
            return (
                <div className={styles.error}>
                    <p>Error loading calendar: {error}</p>
                </div>
            );
        }

        if (loading && layout === 'carousel') {
            return (
                <div className={styles.loading}>
                    <div className={styles.spinner}></div>
                </div>
            );
        }

        return (
            <div className={`${styles.pageCalendar} ${this.props.isDarkTheme ? styles.darkTheme : ''}`}>
                <div className={styles.container}>
                    {layout === 'carousel' ? this.renderCarouselLayout() : this.renderCalendarGridLayout()}
                </div>
            </div>
        );
    }
}
