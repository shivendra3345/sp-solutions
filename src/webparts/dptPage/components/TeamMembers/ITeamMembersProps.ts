import { WebPartContext } from '@microsoft/sp-webpart-base';
import { ITeamMember } from '../../services/TeamMembersService';

export interface ITeamMembersProps {
    context: WebPartContext;
    listTitle: string;
    title?: string;
    department?: string;
    isDarkTheme?: boolean;
    itemsPerRow?: number;
}

export interface ITeamMembersState {
    teamMembers: ITeamMember[];
    loading: boolean;
    error?: string;
}
