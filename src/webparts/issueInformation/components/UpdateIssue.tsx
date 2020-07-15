import * as React from 'react';
import styles from './IssueInformation.module.scss';
import { IIssueInformationProps } from './IIssueInformationProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default class UpdateIssue extends React.Component<IIssueInformationProps, {}> {
  public render(): React.ReactElement<IIssueInformationProps> {
    return (
      <div className={ styles.issueInformation }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Update  Issue!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>
    );
  }
}
