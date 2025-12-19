import * as React from 'react';
import { useEffect, useState } from 'react';
import {
    Persona,
    PersonaSize,
    Dialog,
    DialogFooter,
    PrimaryButton,
    DefaultButton,
    TextField

} from '@fluentui/react';
import styles from './WelcomeNewJoiners.module.scss';
import { SPHttpClient } from '@microsoft/sp-http';

interface IJoinee {
    Id: number;
    Title: string;
    FirstName?: string;
    LastName?: string;
    EmployeeTitle?: string;
    StartDate?: string;
    Department?: string;
    SiteLocation?: string;
    EmployeeIs?: string;
    Status?: string;
    Picture?: string;
}

interface IProps {
    context: any; // WebPartContext
}

const WelcomeNewJoiners: React.FC<IProps> = ({ context }) => {
    const [joinees, setJoinees] = useState<IJoinee[]>([]);
    const [selected, setSelected] = useState<IJoinee | null>(null);
    const [dialogOpen, setDialogOpen] = useState(false);
    const [comment, setComment] = useState('');

    const getField = (obj: any, candidates: string[]) => {
        for (const c of candidates) {
            if (obj[c] !== undefined && obj[c] !== null) return obj[c];
        }
        return undefined;
    };

    useEffect(() => {
        const fetchJoinees = async () => {
            try {
                // Use list GUID provided by user: b08ca367-db69-45f2-8f03-1d4370471de6
                const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists(guid'b08ca367-db69-45f2-8f03-1d4370471de6')/items?$top=6&$orderby=Id desc`;
                const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
                const data = await res.json();

                const items = (data.value || []).map((i: any) => {
                    const first = getField(i, ['Employee_x0020_First_x0020_Name', 'EmployeeFirstName', 'EmployeeFirstName0', 'Employee%20First%20Name', 'EmployeeFirstName0']);
                    const last = getField(i, ['Employee_x0020_Last_x0020_Name', 'EmployeeLastName', 'Employee%20Last%20Name']);
                    const empTitle = getField(i, ['Employee_x0020_Title', 'EmployeeTitle', 'Employee%20Title', 'JobTitle']) || getField(i, ['Title']);
                    const start = getField(i, ['Start_x0020_Date', 'StartDate', 'Start_x0020_Date0']);
                    const dept = getField(i, ['Department']);
                    const site = getField(i, ['Site_x0020_Location', 'SiteLocation']);
                    const employeeIs = getField(i, ['Employee_x0020_Is', 'EmployeeIs', 'Employee%20Is', 'EmployeeIs0']);
                    const status = getField(i, ['Status', 'Status0', 'Status_x0020_']);

                    const displayName = (first || last) ? `${first || ''} ${last || ''}`.trim() : (i.Title || 'New Joiner');

                    return {
                        Id: i.Id,
                        Title: displayName,
                        FirstName: first,
                        LastName: last,
                        EmployeeTitle: empTitle,
                        StartDate: start ? (new Date(start)).toISOString() : undefined,
                        Department: dept,
                        SiteLocation: site,
                        EmployeeIs: employeeIs,
                        Status: status,
                        Picture: getField(i, ['Picture'])
                    } as IJoinee;
                });

                const normalized = (v: any) => (v || '').toString().toLowerCase();
                const filtered = items.filter((it: IJoinee) => {
                    const ei = normalized(it.EmployeeIs);
                    const st = normalized(it.Status);
                    const isNew = ei.includes('new') || ei.includes('new hire');
                    const isComplete = st.includes('complete') || st.includes('completed');
                    return isNew && isComplete;
                });

                setJoinees(filtered.slice(0, 6));
            } catch (e) {
                // If list not found or error, fallback to sample data
                console.warn('Unable to fetch NewJoinees list, using sample data.', e);
                setJoinees([
                    { Id: 1, Title: 'Aisha Khan', EmployeeTitle: 'Software Engineer', EmployeeIs: 'New Hire', Status: 'Complete' },
                ]);
            }
        };

        fetchJoinees();
    }, [context]);

    const openDialog = (p: IJoinee) => {
        setSelected(p);
        setComment('');
        setDialogOpen(true);
    };

    const sendGreeting = async () => {
        if (!selected) return;
        try {
            const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('WelcomeComments')/items`;
            const body = {
                Title: `Greeting for ${selected.Title}`,
                Comment: comment || ''
            } as any;

            await context.spHttpClient.post(url, SPHttpClient.configurations.v1, {
                headers: {
                    'Accept': 'application/json;odata=nometadata',
                    'Content-type': 'application/json;odata=nometadata'
                },
                body: JSON.stringify(body)
            });

            setDialogOpen(false);
            setComment('');
            // give user feedback
            // eslint-disable-next-line no-alert
            alert('Greeting sent.');
        } catch (e) {
            console.error('Failed to post greeting', e);
            // fallback behaviour
            setDialogOpen(false);
            // eslint-disable-next-line no-alert
            alert('Unable to save greeting to SharePoint. (Check list "WelcomeComments")');
        }
    };

    return (
        <div className={styles.welcomePane}>
            <h3 className={styles.title}>Welcome New Joiners</h3>
            <div className={styles.list}>
                {joinees.map(j => (
                    <div key={j.Id} className={styles.joinee} onClick={() => openDialog(j)}>
                        <Persona text={j.Title} secondaryText={j.EmployeeTitle || j.Department} size={PersonaSize.size40} />
                    </div>
                ))}
            </div>

            <Dialog hidden={!dialogOpen} onDismiss={() => setDialogOpen(false)} dialogContentProps={{ title: selected ? `Send greeting to ${selected.Title}` : 'Send greeting' }}>
                <div style={{ minHeight: 80 }}>
                    <TextField label="Message" multiline rows={4} value={comment} onChange={(_, v) => setComment(v || '')} />
                </div>
                <DialogFooter>
                    <PrimaryButton onClick={sendGreeting} text="Send" />
                    <DefaultButton onClick={() => setDialogOpen(false)} text="Cancel" />
                </DialogFooter>
            </Dialog>
        </div>
    );
};

export default WelcomeNewJoiners;
