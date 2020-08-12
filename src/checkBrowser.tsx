import * as React from 'react';
import { escape } from '@microsoft/sp-lodash-subset';

export default class CheckBrowser extends React.Component<{}, {}> {
  public render(): React.ReactElement<{}> {
    return (
        <div>
            <h5>This browser is out of date and may not be compitable with our website, please use Chrome or other modern browsers</h5>
        </div>
    );
  }
}
