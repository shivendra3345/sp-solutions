import * as React from 'react';
import { ITeamMembersProps, ITeamMembersState } from '../TeamMembers/ITeamMembersProps';
import { ITeamMember } from '../../services/TeamMembersService';
import TeamMembersService from '../../services/TeamMembersService';
import styles from './TeamMembers.module.scss';

export class TeamMembers extends React.Component<ITeamMembersProps, ITeamMembersState> {
    private teamMembersService: TeamMembersService;

    constructor(props: ITeamMembersProps) {
        super(props);

        this.state = {
            teamMembers: [],
            loading: true,
            error: undefined
        };

        this.teamMembersService = new TeamMembersService(props.context);
    }

    public componentDidMount(): void {
        this.loadTeamMembers();
    }

    private loadTeamMembers = async (): Promise<void> => {
        try {
            this.setState({ loading: true, error: undefined });

            const listTitle = this.props.listTitle || 'Team Members';
            const members = await this.teamMembersService.getTeamMembers(listTitle);

            this.setState({
                teamMembers: members,
                loading: false
            });
        } catch (error) {
            console.error('Error loading team members:', error);
            this.setState({
                error: 'Failed to load team members. Please try again later.',
                loading: false
            });
        }
    }

    private renderLoading(): JSX.Element {
        return (
            <div className={styles.loading}>
                <div className={styles.spinner}></div>
                <p>Loading team members...</p>
            </div>
        );
    }

    private renderError(): JSX.Element {
        return (
            <div className={styles.error}>
                <p>{this.state.error}</p>
                <button onClick={this.loadTeamMembers} style={{ marginTop: '10px', padding: '8px 16px' }}>
                    Retry
                </button>
            </div>
        );
    }

    private renderNoMembers(): JSX.Element {
        return (
            <div className={styles.noMembers}>
                <p>No team members found. Please configure the Team Members list.</p>
            </div>
        );
    }

    private renderMemberCard(member: ITeamMember): JSX.Element {
        return (
            <div key={member.id} className={styles.memberCard}>
                <div className={styles.memberPhoto}>
                    <img src={member.photoUrl} alt={member.displayName} />
                </div>
                <div className={styles.memberInfo}>
                    <h3 className={styles.memberName}>{member.displayName}</h3>
                    <p className={styles.memberJobTitle}>{member.jobTitle}</p>

                    <div className={styles.memberContact}>
                        {member.email && (
                            <div className={styles.contactItem}>
                                <strong>Email:</strong>
                                <a href={`mailto:${member.email}`} style={{ color: 'inherit', textDecoration: 'none' }}>
                                    {member.email}
                                </a>
                            </div>
                        )}
                        {member.phone && (
                            <div className={styles.contactItem}>
                                <strong>Phone:</strong>
                                <a href={`tel:${member.phone}`} style={{ color: 'inherit', textDecoration: 'none' }}>
                                    {member.phone}
                                </a>
                            </div>
                        )}
                    </div>

                    {member.location && (
                        <div className={styles.memberLocation}>
                            <strong>📍</strong>
                            <span>{member.location}</span>
                        </div>
                    )}
                </div>
            </div>
        );
    }

    private renderMembers(): JSX.Element {
        const members = this.state.teamMembers;

        if (members.length === 0) {
            return this.renderNoMembers();
        }

        return (
            <div className={styles.membersGrid}>
                {members.map((member: ITeamMember) => this.renderMemberCard(member))}
            </div>
        );
    }

    public render(): React.ReactElement<ITeamMembersProps> {
        const { loading, error } = this.state;
        const { title, isDarkTheme } = this.props;

        const containerClassName = `${styles.pageTeamMembers} ${isDarkTheme ? styles.darkTheme : ''}`;

        return (
            <div className={containerClassName}>
                <div className={styles.container}>
                    {title && (
                        <div className={styles.membersHeader}>
                            <h2 className={styles.membersTitle}>{title}</h2>
                        </div>
                    )}

                    {loading && this.renderLoading()}
                    {error && this.renderError()}
                    {!loading && !error && this.renderMembers()}
                </div>
            </div>
        );
    }
}
