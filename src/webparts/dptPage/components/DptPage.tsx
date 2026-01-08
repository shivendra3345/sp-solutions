import * as React from 'react';
import styles from './DptPage.module.scss';
import { IDptPageProps } from './IDptPageProps';
import PageBanner from './PageBanner/PageBanner';
import QuickLinks from './QuickLinks/QuickLinks';
import Calendar from './Calendar/Calendar';
import { TeamMembers } from './TeamMembers/TeamMembers';

export default class DptPage extends React.Component<IDptPageProps, {}> {
  public render(): React.ReactElement<IDptPageProps> {
    return (
      <div className={styles.dptPage}>
        {this.props.showPageBanner && (
          <PageBanner
            context={this.props.context}
            layout={this.props.bannerLayout}
            listTitle={this.props.bannerListTitle}
          />
        )}

        {this.props.showQuickLinks && (
          <QuickLinks
            context={this.props.context}
            listTitle={this.props.quickLinksListTitle || 'QuickLinks'}
            title={this.props.quickLinksTitle}
          />
        )}

        <div className={styles.contentRow}>
          {this.props.showCalendar && (
            <div className={styles.calendarSection}>
              <Calendar
                context={this.props.context}
                layout={this.props.calendarLayout}
                calendarSource={this.props.calendarSource}
                groupId={this.props.groupId}
                spSiteUrl={this.props.spSiteUrl}
                spListTitle={this.props.spListTitle}
              />
            </div>
          )}

          {this.props.showTeamMembers && (
            <div className={styles.teamMembersSection}>
              <TeamMembers
                context={this.props.context}
                listTitle={this.props.teamMembersListTitle || 'Team Members'}
                title={this.props.teamMembersTitle}
                itemsPerRow={this.props.teamMembersItemsPerRow}
              />
            </div>
          )}
        </div>
      </div>
    );
  }
}
